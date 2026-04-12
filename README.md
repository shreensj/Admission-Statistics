# Admission Statistics Dashboard

This project is a static HTML dashboard designed for GitHub Pages with Firebase as the backend.
It starts from the Excel snapshot generated from:

`D:\Niraj Projects\Excel\Admission Statistics from Jan 2026.xlsx`

## Main files

- `index.html` is the dashboard page.
- `assets/app.js` loads the workbook snapshot, connects to Firebase, reads Firestore data, and saves new entries.
- `firebase-config.js` is where you paste your Firebase web app config.
- `firestore.rules` contains a starter rules file for Cloud Firestore.
- `data/workbook-data.js` is the extracted workbook snapshot used for first load and initial Firebase seeding.
- `scripts/extract_excel.ps1` refreshes the workbook snapshot from the Excel file.
- `refresh-dashboard-data.cmd` is a quick shortcut to regenerate `data/workbook-data.js`.

## How the app works

- The dashboard uses `data/workbook-data.js` as the starter schema and local fallback.
- If Firebase is configured, the app signs in anonymously and reads rows from Firestore.
- Guests can view data.
- Admin users can seed Firebase, add rows, edit rows, and delete rows.
- The `Seed Firebase` button uploads the workbook snapshot into Firestore the first time.

## Firebase setup steps

1. Go to the [Firebase Console](https://console.firebase.google.com/).
2. Create a new Firebase project.
3. In the project, create a Web app.
4. Copy the Firebase web config and paste it into `firebase-config.js`.
5. In Firebase Console, enable `Authentication`.
6. Inside Authentication, enable `Anonymous` sign-in.
7. Also enable `Email/Password` sign-in.
8. In Authentication, create your admin user email and password.
9. Sign in once in the app using that email and password so Firebase creates the user account and UID.
10. In Firebase Console, create a `Cloud Firestore` database in production or test mode.
11. Open the Rules tab in Firestore and replace the rules with the content of `firestore.rules`.
12. Publish the rules.
13. In Firestore Data, create a collection named `admins`.
14. Inside `admins`, create a document whose document ID is your admin user's Firebase UID.
15. Add a field like `email` with your admin email address.
16. Open your dashboard, log in as admin, click `Seed Firebase`, and wait for the workbook data to upload.

## How to get the admin UID

1. Open the app after Firebase is connected.
2. Click `Admin login`.
3. Log in with your Firebase email and password.
4. After login, copy the UID shown in the `Signed in user` area if needed.
5. Use that UID as the Firestore document ID inside the `admins` collection.

## Permissions model

- Anonymous users can read dashboard data.
- Only users whose UID exists in Firestore collection `admins` can write data.
- Write actions include seed, add, edit, and delete.

## GitHub upload steps

1. Create a new GitHub repository.
2. Upload the project files to the repository root.
3. Make sure `index.html` and `.nojekyll` stay in the root of the repository.
4. Commit and push the files.
5. In GitHub, open the repository `Settings`.
6. Open `Pages`.
7. Under build and deployment, choose `Deploy from a branch`.
8. Select your main branch and the `/ (root)` folder.
9. Save the settings and wait for GitHub Pages to publish your site.
10. Open the published GitHub Pages URL and test Firebase read and write behavior.

## Local testing note

- Because this app uses JavaScript modules and Firebase CDN imports, testing through GitHub Pages is the easiest path.
- If you test locally and the browser blocks module loading from `file://`, use a small local web server instead of double-click opening the file.

## Updating from Excel later

1. Edit the Excel file.
2. Run `refresh-dashboard-data.cmd`.
3. Commit the updated `data/workbook-data.js` file to GitHub.
4. If you want Firestore to match the new snapshot, clear old rows manually or seed into a new Firebase project or new dashboard path.

## Notes

- Firebase web config is safe to include in a frontend app, but your Firestore rules must be correct.
- This project uses anonymous sign-in for viewing and email/password admin login for changes.
- If you want a stricter setup later, the next upgrade should be custom admin claims or a backend API.
