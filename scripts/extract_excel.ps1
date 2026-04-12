param(
  [string]$WorkbookPath = 'D:\Niraj Projects\Excel\Admission Statistics from Jan 2026.xlsx',
  [string]$OutputPath = 'C:\Users\user\Documents\New project\data\workbook-data.js'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ColIndex {
  param([string]$CellRef)

  $letters = ($CellRef -replace '\d', '').ToUpperInvariant()
  $sum = 0
  foreach ($ch in $letters.ToCharArray()) {
    $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $sum
}

function Convert-ExcelSerialToDate {
  param([double]$Value)

  if ($Value -lt 1) {
    return $Value.ToString()
  }

  return ([datetime]'1899-12-30').AddDays($Value).ToString('yyyy-MM-dd')
}

function Normalize-CellValue {
  param(
    [string]$Header,
    [string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ''
  }

  if ($Header -match 'Date|Birth|Mth') {
    $parsed = 0.0
    if ([double]::TryParse($Value, [ref]$parsed) -and $parsed -gt 20000) {
      return Convert-ExcelSerialToDate -Value $parsed
    }
  }

  return $Value.Trim()
}

function Get-UniqueHeaders {
  param([string[]]$Headers)

  $seen = @{}
  $result = @()

  foreach ($header in $Headers) {
    $baseName = if ([string]::IsNullOrWhiteSpace($header)) { 'Column' } else { $header.Trim() }
    if (-not $seen.ContainsKey($baseName)) {
      $seen[$baseName] = 1
      $result += $baseName
      continue
    }

    $seen[$baseName] += 1
    $result += "$baseName ($($seen[$baseName]))"
  }

  return $result
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
  throw "Workbook not found: $WorkbookPath"
}

$outputDir = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)

try {
  function Get-EntryText {
    param([string]$FullName)

    $entry = $zip.GetEntry($FullName)
    if (-not $entry) {
      return $null
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
      return $reader.ReadToEnd()
    }
    finally {
      $reader.Dispose()
    }
  }

  $sharedStrings = @()
  $sharedText = Get-EntryText -FullName 'xl/sharedStrings.xml'
  if ($sharedText) {
    $sharedDoc = [xml]$sharedText
    $sharedNs = [System.Xml.XmlNamespaceManager]::new($sharedDoc.NameTable)
    $sharedNs.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    foreach ($si in $sharedDoc.SelectNodes('//d:si', $sharedNs)) {
      $parts = @()
      foreach ($textNode in $si.SelectNodes('.//d:t', $sharedNs)) {
        $parts += $textNode.InnerText
      }
      $sharedStrings += ($parts -join '')
    }
  }

  $workbookDoc = [xml](Get-EntryText -FullName 'xl/workbook.xml')
  $workbookNs = [System.Xml.XmlNamespaceManager]::new($workbookDoc.NameTable)
  $workbookNs.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
  $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

  $relsDoc = [xml](Get-EntryText -FullName 'xl/_rels/workbook.xml.rels')
  $relsNs = [System.Xml.XmlNamespaceManager]::new($relsDoc.NameTable)
  $relsNs.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
  $relationshipMap = @{}
  foreach ($rel in $relsDoc.SelectNodes('//p:Relationship', $relsNs)) {
    $relationshipMap[$rel.Id] = $rel.Target
  }

  $sheetPayload = @()

  foreach ($sheet in $workbookDoc.SelectNodes('//d:sheets/d:sheet', $workbookNs)) {
    $sheetName = $sheet.GetAttribute('name')
    $relationshipId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $sheetPath = 'xl/' + $relationshipMap[$relationshipId]
    $sheetDoc = [xml](Get-EntryText -FullName $sheetPath)
    $sheetNs = [System.Xml.XmlNamespaceManager]::new($sheetDoc.NameTable)
    $sheetNs.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $rows = @($sheetDoc.SelectNodes('//d:sheetData/d:row', $sheetNs))
    if ($rows.Count -eq 0) {
      $sheetPayload += [ordered]@{
        name = $sheetName
        headers = @()
        rows = @()
      }
      continue
    }

    $rowMaps = @()
    $maxColumns = 0

    foreach ($row in $rows) {
      $valueMap = @{}
      foreach ($cell in $row.SelectNodes('./d:c', $sheetNs)) {
        $ref = $cell.GetAttribute('r')
        $colIndex = Get-ColIndex -CellRef $ref
        $cellType = $cell.GetAttribute('t')
        $valueNode = $cell.SelectSingleNode('./d:v', $sheetNs)
        $inlineNode = $cell.SelectSingleNode('./d:is/d:t', $sheetNs)

        $value = ''
        if ($cellType -eq 's' -and $valueNode) {
          $value = $sharedStrings[[int]$valueNode.InnerText]
        }
        elseif ($cellType -eq 'inlineStr' -and $inlineNode) {
          $value = $inlineNode.InnerText
        }
        elseif ($valueNode) {
          $value = $valueNode.InnerText
        }

        $valueMap[$colIndex] = $value
        if ($colIndex -gt $maxColumns) {
          $maxColumns = $colIndex
        }
      }
      $rowMaps += ,$valueMap
    }

    $headers = @()
    for ($i = 1; $i -le $maxColumns; $i++) {
      $header = ''
      if ($rowMaps[0].ContainsKey($i)) {
        $header = [string]$rowMaps[0][$i]
      }
      if ([string]::IsNullOrWhiteSpace($header)) {
        $header = "Column $i"
      }
      $headers += $header.Trim()
    }
    $headers = Get-UniqueHeaders -Headers $headers

    $dataRows = @()
    for ($r = 1; $r -lt $rowMaps.Count; $r++) {
      $record = [ordered]@{}
      for ($c = 1; $c -le $headers.Count; $c++) {
        $rawValue = ''
        if ($rowMaps[$r].ContainsKey($c)) {
          $rawValue = [string]$rowMaps[$r][$c]
        }
        $record[$headers[$c - 1]] = Normalize-CellValue -Header $headers[$c - 1] -Value $rawValue
      }
      $dataRows += ,$record
    }

    $sheetPayload += [ordered]@{
      name = $sheetName
      headers = $headers
      rows = $dataRows
    }
  }

  $payload = [ordered]@{
    sourceWorkbook = $WorkbookPath
    generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    sheets = $sheetPayload
  }

  $json = $payload | ConvertTo-Json -Depth 8
  $fileContent = "window.WORKBOOK_DATA = $json;"
  Set-Content -LiteralPath $OutputPath -Value $fileContent -Encoding UTF8

  Write-Output "Dashboard data written to $OutputPath"
}
finally {
  $zip.Dispose()
}
