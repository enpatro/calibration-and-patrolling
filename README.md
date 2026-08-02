# Compliance Portal Firebase + GitHub Pages Bundle v5

This is the complete hosted portal bundle for Calibration Management and Patrolling Management.

## Included
- GitHub Pages frontend
- Firebase Auth login
- Firestore live database sync
- Firebase Storage certificate PDF upload
- Local JSON fallback if Firebase is not configured
- Calibration and patrolling Excel templates
- CSV templates
- Firestore and Storage security rule files

## Quick Hosting on GitHub Pages
1. Extract this zip.
2. Upload all files to GitHub repository root.
3. Open Settings > Pages.
4. Select Deploy from branch > main > root.
5. Open the GitHub Pages URL.

## Firebase Setup
1. Go to Firebase Console and create a project.
2. Add a Web App.
3. Copy Firebase web config.
4. Paste config into assets/js/firebase-config.js.
5. Enable Authentication > Email/Password.
6. Create your admin user in Firebase Authentication.
7. Create Firestore Database.
8. Create Storage.
9. Add GitHub Pages URL in Authentication > Settings > Authorized domains.
10. Upload firebase/firestore.rules and firebase/storage.rules rules manually in Firebase Console, or use Firebase CLI.

## Firestore Collections
- calibration
- patrolling
- auditLogs
- users

## Bulk Upload to Firestore
1. Login as admin in Bulk/Admin tab.
2. Upload calibration CSV/JSON.
3. Click Sync Calibration to Firestore.
4. Upload patrolling CSV/JSON.
5. Click Sync Patrolling to Firestore.
6. Upload certificate PDFs with Upload Certificates.

## Important
If Firebase is not configured or user is not logged in, portal automatically uses data/calibration.json and data/patrolling.json.

## Changes in v6
- Added missing dashboard quick access tab cards below the dashboard heading.
- Dashboard now has visible cards for Calibration, Patrolling, Open NG, Bulk/Admin and Firebase Setup.
- Cards are clickable and navigate to respective sections.
