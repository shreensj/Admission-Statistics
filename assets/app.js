import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signInAnonymously,
  onAuthStateChanged,
  signOut,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  addDoc,
  collection,
  deleteDoc,
  doc,
  getDoc,
  getDocs,
  getFirestore,
  limit,
  query,
  serverTimestamp,
  setDoc,
  updateDoc,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const DASHBOARD_ID = "admission-statistics";
const baseData = window.WORKBOOK_DATA || { sheets: [] };
const firebaseConfig = window.FIREBASE_CONFIG;

const state = {
  workbook: cloneData(baseData),
  activeSheetIndex: 0,
  searchTerm: "",
  db: null,
  auth: null,
  firebaseReady: false,
  user: null,
  usingFirebaseData: false,
  isAdmin: false,
  editingRowId: null,
};

const summaryGrid = document.getElementById("summaryGrid");
const sheetTabs = document.getElementById("sheetTabs");
const activeSheetTitle = document.getElementById("activeSheetTitle");
const rowCountValue = document.getElementById("rowCountValue");
const columnCountValue = document.getElementById("columnCountValue");
const sourceValue = document.getElementById("sourceValue");
const sheetTable = document.getElementById("sheetTable");
const searchInput = document.getElementById("searchInput");
const topDate = document.getElementById("topDate");
const overviewTabBtn = document.getElementById("overviewTabBtn");
const sheetCountBadge = document.getElementById("sheetCountBadge");
const recordCountChip = document.getElementById("recordCountChip");
const viewPendingBtn = document.getElementById("viewPendingBtn");
const entryDialog = document.getElementById("entryDialog");
const entryForm = document.getElementById("entryForm");
const formFields = document.getElementById("formFields");
const dialogTitle = document.getElementById("dialogTitle");
const entryDialogLabel = document.getElementById("entryDialogLabel");
const exportJsonBtn = document.getElementById("exportJsonBtn");
const exportSheetBtn = document.getElementById("exportSheetBtn");
const addEntryBtn = document.getElementById("addEntryBtn");
const closeDialogBtn = document.getElementById("closeDialogBtn");
const cancelDialogBtn = document.getElementById("cancelDialogBtn");
const seedFirebaseBtn = document.getElementById("seedFirebaseBtn");
const adminLoginBtn = document.getElementById("adminLoginBtn");
const adminLogoutBtn = document.getElementById("adminLogoutBtn");
const adminStatusValue = document.getElementById("adminStatusValue");
const firebaseStatusValue = document.getElementById("firebaseStatusValue");
const authStatusValue = document.getElementById("authStatusValue");
const dataSourceValue = document.getElementById("dataSourceValue");
const statusMessageValue = document.getElementById("statusMessageValue");
const adminDialog = document.getElementById("adminDialog");
const adminForm = document.getElementById("adminForm");
const closeAdminDialogBtn = document.getElementById("closeAdminDialogBtn");
const cancelAdminDialogBtn = document.getElementById("cancelAdminDialogBtn");
const pendingDialog = document.getElementById("pendingDialog");
const pendingList = document.getElementById("pendingList");
const pendingStudentCount = document.getElementById("pendingStudentCount");
const pendingAmountTotal = document.getElementById("pendingAmountTotal");
const closePendingDialogBtn = document.getElementById("closePendingDialogBtn");
const closePendingFooterBtn = document.getElementById("closePendingFooterBtn");

boot();

async function boot() {
  if (topDate) {
    topDate.textContent = new Date().toLocaleDateString("en-IN", {
      weekday: "short",
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  }
  render();
  if (window.location.protocol === "file:") {
    setStatus("This dashboard should be opened through GitHub Pages or a local web server, not directly as a file.", "warn");
  }
  if (!hasFirebaseConfig()) {
    setStatus("Firebase config file is still a placeholder. Update firebase-config.js to connect Firestore.", "warn");
    return;
  }

  try {
    const app = initializeApp(firebaseConfig);
    state.db = getFirestore(app);
    state.auth = getAuth(app);
    state.firebaseReady = true;
    firebaseStatusValue.textContent = "Configured";
    setStatus("Firebase initialized. Signing in anonymously and loading cloud data...");

    onAuthStateChanged(state.auth, async (user) => {
      state.user = user || null;
      authStatusValue.textContent = user?.email || user?.uid || "Signed out";
      state.isAdmin = user ? await checkAdminAccess(user) : false;
      adminStatusValue.textContent = !user ? "Guest" : user.isAnonymous ? "Guest" : state.isAdmin ? "Admin" : "No access";
      adminLoginBtn.classList.toggle("hidden", Boolean(user && !user.isAnonymous));
      adminLogoutBtn.classList.toggle("hidden", !(user && !user.isAnonymous));
      if (user) {
        await syncFromFirestore();
      }
      render();
    });

    await signInAnonymously(state.auth);
  } catch (error) {
    console.error(error);
    setStatus(`Firebase setup failed: ${error.message}`, "warn");
    firebaseStatusValue.textContent = "Error";
    authStatusValue.textContent = "Unavailable";
  }
}

async function checkAdminAccess(user) {
  if (!user || user.isAnonymous || !state.db) {
    return false;
  }

  try {
    const adminDoc = await getDoc(doc(state.db, "admins", user.uid));
    if (!adminDoc.exists()) {
      setStatus("Logged in, but this account is not in Firestore admins yet.", "warn");
      return false;
    }
    return true;
  } catch (error) {
    console.error(error);
    setStatus(`Could not verify admin access: ${error.message}`, "warn");
    return false;
  }
}

function hasFirebaseConfig() {
  return Boolean(firebaseConfig && firebaseConfig.apiKey && firebaseConfig.projectId && firebaseConfig.appId);
}

function cloneData(value) {
  const cloned = JSON.parse(JSON.stringify(value));
  return normalizeWorkbook(cloned);
}

function getActiveSheet() {
  return state.workbook.sheets[state.activeSheetIndex];
}

function normalizeWorkbook(workbook) {
  workbook.sheets = (workbook.sheets || []).map((sheet) => ({
    ...sheet,
    rows: (sheet.rows || []).map((row, index) => normalizeRow(row, `${slugify(sheet.name)}-local-${index}`)),
  }));
  return workbook;
}

function normalizeRow(row, fallbackId) {
  if (row && typeof row === "object" && "data" in row) {
    return {
      id: row.id || fallbackId,
      data: row.data || {},
    };
  }

  return {
    id: fallbackId,
    data: row || {},
  };
}

function getFilteredRows(sheet) {
  if (!state.searchTerm) {
    return sheet.rows;
  }

  const term = state.searchTerm.toLowerCase();
  return sheet.rows.filter((row) =>
    Object.values(row.data || {}).some((value) => String(value || "").toLowerCase().includes(term))
  );
}

function buildSummaryCards() {
  const totalSheets = state.workbook.sheets.length;
  const totalRows = state.workbook.sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
  const totalColumns = state.workbook.sheets.reduce((sum, sheet) => sum + sheet.headers.length, 0);
  const pendingItems = getPendingFeeItems();
  const totalPending = pendingItems.reduce((sum, item) => sum + item.pendingAmount, 0);

  const cards = [
    { label: "Active sheets", value: totalSheets },
    { label: "Total records", value: totalRows },
    { label: "Pending students", value: pendingItems.length },
    { label: "Pending fees", value: formatCurrency(totalPending) },
  ];

  summaryGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="summary-card">
          <span>${escapeHtml(card.label)}</span>
          <strong>${escapeHtml(String(card.value))}</strong>
        </article>
      `
    )
    .join("");
}

function buildSheetTabs() {
  sheetTabs.innerHTML = state.workbook.sheets
    .map((sheet, index) => {
      const isActive = index === state.activeSheetIndex ? "active" : "";
      return `
        <button class="sheet-tab ${isActive}" data-sheet-index="${index}">
          <strong>${escapeHtml(sheet.name)}</strong>
          <span class="nav-badge">${sheet.rows.length}</span>
        </button>
      `;
    })
    .join("");

  if (sheetCountBadge) {
    sheetCountBadge.textContent = String(state.workbook.sheets.length);
  }

  sheetTabs.querySelectorAll("[data-sheet-index]").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeSheetIndex = Number(button.dataset.sheetIndex);
      state.searchTerm = "";
      searchInput.value = "";
      render();
    });
  });
}

function buildTable() {
  const sheet = getActiveSheet();
  const rows = getFilteredRows(sheet);

  activeSheetTitle.textContent = sheet.name;
  dialogTitle.textContent = `Add entry to ${sheet.name}`;
  rowCountValue.textContent = rows.length;
  columnCountValue.textContent = sheet.headers.length;
  sourceValue.textContent = state.workbook.sourceWorkbook || baseData.sourceWorkbook || "-";
  if (recordCountChip) {
    recordCountChip.textContent = `${rows.length} rows`;
  }

  const thead = `
    <thead>
      <tr>
        ${sheet.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}
        ${state.isAdmin ? "<th>Actions</th>" : ""}
      </tr>
    </thead>
  `;

  const tbody = rows.length
    ? `
      <tbody>
        ${rows
          .map(
            (row) => `
              <tr>
                ${sheet.headers.map((header) => `<td>${escapeHtml(String(row.data?.[header] ?? ""))}</td>`).join("")}
                ${
                  state.isAdmin
                    ? `<td>
                        <div class="row-actions">
                          <button class="row-action" data-edit-row="${row.id}">Edit</button>
                          <button class="row-action danger" data-delete-row="${row.id}">Delete</button>
                        </div>
                      </td>`
                    : ""
                }
              </tr>
            `
          )
          .join("")}
      </tbody>
    `
    : '<tbody><tr><td class="empty-state" colspan="100%">No matching rows found for this sheet.</td></tr></tbody>';

  sheetTable.innerHTML = thead + tbody;

  if (state.isAdmin) {
    sheetTable.querySelectorAll("[data-edit-row]").forEach((button) => {
      button.addEventListener("click", () => openEntryDialog(button.dataset.editRow));
    });
    sheetTable.querySelectorAll("[data-delete-row]").forEach((button) => {
      button.addEventListener("click", () => deleteRow(button.dataset.deleteRow));
    });
  }
}

function getPendingFeeItems() {
  const mainSheet = state.workbook.sheets.find((sheet) => sheet.name.toLowerCase() === "main");
  if (!mainSheet) {
    return [];
  }

  return mainSheet.rows
    .map((row) => {
      const data = row.data || {};
      const studentName = data["Student Name"] || "";
      const rollNo = data["Roll No (2)"] || data["Roll No"] || "";
      const courseName = data["Course Name"] || data["First Course"] || data["Course Type"] || "";
      const totalFees = parseNumber(data["Total Fees"]);
      const pending = parseNumber(data["Fees Pending"]);
      return {
        id: row.id,
        studentName,
        rollNo,
        courseName,
        totalFees,
        pendingAmount: pending,
      };
    })
    .filter((item) => item.studentName && item.pendingAmount > 0)
    .sort((a, b) => b.pendingAmount - a.pendingAmount);
}

function parseNumber(value) {
  const normalized = String(value ?? "").replace(/[^0-9.-]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

function buildPendingPopup() {
  if (!pendingList || !pendingStudentCount || !pendingAmountTotal) {
    return;
  }

  const items = getPendingFeeItems();
  pendingStudentCount.textContent = String(items.length);
  pendingAmountTotal.textContent = formatCurrency(items.reduce((sum, item) => sum + item.pendingAmount, 0));

  if (!items.length) {
    pendingList.innerHTML = '<div class="empty-state">No pending fees found in the current workbook data.</div>';
    return;
  }

  pendingList.innerHTML = items
    .map((item) => {
      const initials = item.studentName
        .split(" ")
        .filter(Boolean)
        .slice(0, 2)
        .map((part) => part[0]?.toUpperCase() || "")
        .join("");
      return `
        <article class="pending-item">
          <div class="pending-avatar">${escapeHtml(initials || "ST")}</div>
          <div class="pending-meta">
            <strong>${escapeHtml(item.studentName)}</strong>
            <span>${escapeHtml(item.rollNo)} | ${escapeHtml(item.courseName || "Course not set")}</span>
          </div>
          <div class="pending-amount">
            <strong>${escapeHtml(formatCurrency(item.pendingAmount))}</strong>
            <span>Total fees ${escapeHtml(formatCurrency(item.totalFees))}</span>
          </div>
        </article>
      `;
    })
    .join("");
}

function openEntryDialog(rowId = null) {
  const sheet = getActiveSheet();
  const existingRow = rowId ? sheet.rows.find((row) => row.id === rowId) : null;
  state.editingRowId = rowId;
  entryDialogLabel.textContent = rowId ? "Edit row" : "New row";
  formFields.innerHTML = sheet.headers
    .map((header, index) => {
      const inputType = /date|birth/i.test(header) ? "date" : "text";
      const fieldId = `${slugify(header) || "field"}-${index}`;
      const value = existingRow?.data?.[header] ?? "";
      return `
        <div class="form-field">
          <label for="${fieldId}">${escapeHtml(header)}</label>
          <input id="${fieldId}" name="${escapeHtmlAttribute(header)}" type="${inputType}" value="${escapeHtmlAttribute(String(value))}" />
        </div>
      `;
    })
      .join("");
  dialogTitle.textContent = rowId ? `Edit entry in ${sheet.name}` : `Add entry to ${sheet.name}`;
  entryDialog.showModal();
}

async function addNewRow(formData) {
  const sheet = getActiveSheet();
  const row = {};
  sheet.headers.forEach((header) => {
    row[header] = (formData.get(header) || "").toString().trim();
  });

  if (!state.firebaseReady || !state.user || !state.isAdmin) {
    setStatus("Admin login is required to add or edit entries.", "warn");
    return false;
  }

  try {
    const rowsRef = collection(state.db, "dashboards", DASHBOARD_ID, "sheets", slugify(sheet.name), "rows");
    if (state.editingRowId) {
      await updateDoc(doc(rowsRef, state.editingRowId), {
        data: row,
        updatedAt: serverTimestamp(),
      });
      sheet.rows = sheet.rows.map((item) =>
        item.id === state.editingRowId ? { ...item, data: row } : item
      );
      setStatus(`Entry updated in ${sheet.name}.`);
    } else {
      const created = await addDoc(rowsRef, {
        data: row,
        createdAt: serverTimestamp(),
        sheetName: sheet.name,
      });
      sheet.rows.unshift({ id: created.id, data: row });
      setStatus(`Entry saved to Firebase in ${sheet.name}.`);
    }

    state.usingFirebaseData = true;
    state.editingRowId = null;
    render();
    return true;
  } catch (error) {
    console.error(error);
    setStatus(`Unable to save entry: ${error.message}`, "warn");
    return false;
  }
}

async function deleteRow(rowId) {
  if (!state.isAdmin) {
    setStatus("Admin login is required to delete entries.", "warn");
    return;
  }

  const confirmed = window.confirm("Delete this row permanently?");
  if (!confirmed) {
    return;
  }

  const sheet = getActiveSheet();

  try {
    const rowRef = doc(state.db, "dashboards", DASHBOARD_ID, "sheets", slugify(sheet.name), "rows", rowId);
    await deleteDoc(rowRef);
    sheet.rows = sheet.rows.filter((row) => row.id !== rowId);
    setStatus(`Entry deleted from ${sheet.name}.`);
    render();
  } catch (error) {
    console.error(error);
    setStatus(`Unable to delete entry: ${error.message}`, "warn");
  }
}

async function syncFromFirestore() {
  if (!state.db) {
    return;
  }

  try {
    const workbook = cloneData(baseData);
    let foundCloudRows = false;

    for (const sheet of workbook.sheets) {
      const rowsRef = collection(state.db, "dashboards", DASHBOARD_ID, "sheets", slugify(sheet.name), "rows");
      const snapshot = await getDocs(rowsRef);
      if (!snapshot.empty) {
        foundCloudRows = true;
        sheet.rows = snapshot.docs
          .map((docSnap) => ({
            id: docSnap.id,
            data: docSnap.data().data || {},
          }))
          .filter((row) => row.data);
      }
    }

    state.workbook = workbook;
    state.usingFirebaseData = foundCloudRows;
    dataSourceValue.textContent = foundCloudRows ? "Firebase + workbook schema" : "Workbook snapshot";
    setStatus(foundCloudRows ? "Cloud data loaded from Firestore." : "No Firestore rows found yet. You can seed Firebase from the workbook snapshot.");
    render();
  } catch (error) {
    console.error(error);
    setStatus(`Unable to read Firestore data: ${error.message}`, "warn");
  }
}

async function seedFirebaseFromWorkbook() {
  if (!state.firebaseReady || !state.user || !state.isAdmin) {
    setStatus("Admin login is required before seeding Firebase.", "warn");
    return;
  }

  seedFirebaseBtn.disabled = true;
  seedFirebaseBtn.textContent = "Seeding...";

  try {
    for (const sheet of baseData.sheets) {
      const sheetId = slugify(sheet.name);
      const sheetMetaRef = doc(state.db, "dashboards", DASHBOARD_ID, "sheets", sheetId);
      await setDoc(
        sheetMetaRef,
        {
          name: sheet.name,
          headers: sheet.headers,
          seededAt: serverTimestamp(),
        },
        { merge: true }
      );

      const existingRows = await getDocs(query(collection(state.db, "dashboards", DASHBOARD_ID, "sheets", sheetId, "rows"), limit(1)));
      if (!existingRows.empty) {
        continue;
      }

      for (const row of sheet.rows) {
        await addDoc(collection(state.db, "dashboards", DASHBOARD_ID, "sheets", sheetId, "rows"), {
          data: row.data || row,
          createdAt: serverTimestamp(),
          sheetName: sheet.name,
          importedFromWorkbook: true,
        });
      }
    }

    await syncFromFirestore();
    setStatus("Workbook snapshot uploaded to Firebase.");
  } catch (error) {
    console.error(error);
    setStatus(`Seeding failed: ${error.message}`, "warn");
  } finally {
    seedFirebaseBtn.disabled = false;
    seedFirebaseBtn.textContent = "Seed Firebase";
  }
}

function exportJson() {
  downloadFile("dashboard-data.json", JSON.stringify(state.workbook, null, 2), "application/json");
}

function exportCurrentSheetCsv() {
  const sheet = getActiveSheet();
  const lines = [
    sheet.headers.map(csvEscape).join(","),
    ...sheet.rows.map((row) => sheet.headers.map((header) => csvEscape(row.data?.[header] || "")).join(",")),
  ];
  downloadFile(`${slugify(sheet.name)}.csv`, lines.join("\n"), "text/csv");
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function slugify(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeHtmlAttribute(value) {
  return escapeHtml(value);
}

function setStatus(message, tone = "ok") {
  statusMessageValue.textContent = message;
  statusMessageValue.style.color = tone === "warn" ? "var(--red)" : "var(--muted)";
  statusMessageValue.style.fontWeight = tone === "warn" ? "700" : "400";
  firebaseStatusValue.textContent = state.firebaseReady ? "Configured" : "Not connected";
  dataSourceValue.textContent = state.usingFirebaseData ? "Firestore" : "Workbook snapshot";
}

async function loginAdmin(formData) {
  if (!state.firebaseReady || !state.auth) {
    setStatus("Firebase is not ready yet.", "warn");
    return false;
  }

  const email = (formData.get("email") || "").toString().trim();
  const password = (formData.get("password") || "").toString();

  try {
    await signInWithEmailAndPassword(state.auth, email, password);
    adminForm.reset();
    setStatus("Admin login successful.");
    return true;
  } catch (error) {
    console.error(error);
    setStatus(`Admin login failed: ${error.message}`, "warn");
    return false;
  }
}

async function logoutAdmin() {
  if (!state.auth) {
    return;
  }

  try {
    await signOut(state.auth);
    await signInAnonymously(state.auth);
    setStatus("Switched back to guest mode.");
  } catch (error) {
    console.error(error);
    setStatus(`Logout failed: ${error.message}`, "warn");
  }
}

function render() {
  buildSummaryCards();
  buildSheetTabs();
  buildTable();
  buildPendingPopup();
  addEntryBtn.disabled = !state.isAdmin;
  seedFirebaseBtn.disabled = !state.isAdmin;
}

searchInput.addEventListener("input", (event) => {
  state.searchTerm = event.target.value.trim();
  buildTable();
});

addEntryBtn.addEventListener("click", () => openEntryDialog());
closeDialogBtn.addEventListener("click", () => {
  state.editingRowId = null;
  entryDialog.close();
});
cancelDialogBtn.addEventListener("click", () => {
  state.editingRowId = null;
  entryDialog.close();
});
seedFirebaseBtn.addEventListener("click", seedFirebaseFromWorkbook);
adminLoginBtn.addEventListener("click", () => adminDialog.showModal());
adminLogoutBtn.addEventListener("click", logoutAdmin);
closeAdminDialogBtn.addEventListener("click", () => adminDialog.close());
cancelAdminDialogBtn.addEventListener("click", () => adminDialog.close());

entryForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(entryForm);
  const saved = await addNewRow(formData);
  if (saved) {
    entryForm.reset();
    entryDialog.close();
  }
});

adminForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(adminForm);
  const loggedIn = await loginAdmin(formData);
  if (loggedIn) {
    adminDialog.close();
  }
});

exportJsonBtn.addEventListener("click", exportJson);
exportSheetBtn.addEventListener("click", exportCurrentSheetCsv);
if (overviewTabBtn) {
  overviewTabBtn.addEventListener("click", () => {
    searchInput.focus();
  });
}
if (viewPendingBtn) {
  viewPendingBtn.addEventListener("click", () => {
    buildPendingPopup();
    pendingDialog.showModal();
  });
}
if (closePendingDialogBtn) {
  closePendingDialogBtn.addEventListener("click", () => pendingDialog.close());
}
if (closePendingFooterBtn) {
  closePendingFooterBtn.addEventListener("click", () => pendingDialog.close());
}
