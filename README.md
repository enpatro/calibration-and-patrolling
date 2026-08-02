# Compliance Portal V12 Firebase Bulk Write

## What changed
- Bulk calibration upload writes to Firestore collection `calibration`.
- Bulk patrolling upload writes to Firestore collection `patrolling`.
- Auto Write checkbox writes immediately after upload.
- Manual buttons also available: Write Calibration to Firestore, Write Patrolling to Firestore.
- Certificate PDFs upload to Firebase Storage path `certificates/fileName.pdf`.

## Required before Firestore write
1. Create Firebase project.
2. Enable Authentication > Email/Password.
3. Create admin user.
4. Enable Firestore.
5. Enable Storage.
6. Copy Firebase web app config into `assets/js/firebase-config.js`.
7. Upload all files to GitHub root.
8. Login from Bulk/Admin tab.


## V13 Fix
Tabs and pages fixed. Firebase SDK is now loaded dynamically only when Login is clicked, so dashboard navigation will work even before Firebase config is filled.
