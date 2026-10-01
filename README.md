# Lab School Video Behavior Database React Clone

This is a React + Vite clone of the Lab School app, built to run as a static site on GitHub Pages.

## Live site

- [https://ryanlay.github.io/tcfdlabschool.github.io/](https://ryanlay.github.io/tcfdlabschool.github.io/)

## Deployment

- Repository: `https://github.com/ryanlay/tcfdlabschool.github.io`
- GitHub Actions workflow: `.github/workflows/deploy.yml`
- Deploy trigger: push to `main`

## Run locally

```powershell
cd "c:\Users\rlay\OneDrive - The Center For Discovery\Projects\Active\lab-school-database-react"
npm install
npm run dev
```

## Build for GitHub Pages

```powershell
npm run build
```

The app uses `base: '/tcfdlabschool.github.io/'` for project-site hosting on GitHub Pages.

## Features

- Intake wizard with subjects and behavior selection, with second-precision recording start times
- Query-style review tables
- Searchable data view
- Subject Profile: historical review of subjects with Person ID, date of birth, Lab School start/end dates, and a validated set of 20 target behaviors — includes a pre-seeded historical roster (subject codes `S01`-`S23`, `AS01`-`AS15`, `RS01`-`RS04`) with known Person IDs, additively merged in on load without overwriting existing data
- Admin lists for subjects and behaviors
- Export/Import JSON backup tool (Admin ▸ Data Safety) as a manual backup/restore option alongside Firebase
- Shared persistent data in Firebase Firestore
- SharePoint folder links based on the recording date

## Shared Database Setup (Firebase Firestore)

This app stores `subjects`, `behaviors`, `videos`, and an `updatedAt` value in one Firestore document, so data is shared across browsers and devices. The document path is `lab_school_state/shared` by default. The app creates it on first load if it does not exist.

### 1) Configure Firebase

Create a Firebase project, register a web app, and enable Cloud Firestore in Native mode. Copy the web app configuration values into `.env.local` in the project root:

```env
VITE_FIREBASE_API_KEY=your-api-key
VITE_FIREBASE_AUTH_DOMAIN=your-project-id.firebaseapp.com
VITE_FIREBASE_PROJECT_ID=your-project-id
VITE_FIREBASE_STORAGE_BUCKET=your-project-id.firebasestorage.app
VITE_FIREBASE_MESSAGING_SENDER_ID=your-sender-id
VITE_FIREBASE_APP_ID=your-app-id
# Optional; defaults to lab_school_state
VITE_FIREBASE_COLLECTION=lab_school_state
```

Restart the Vite development server after changing environment variables. All six Firebase web app configuration values above must be set for shared storage to be enabled. These are client configuration values, not server credentials.

### 2) Firestore document and security rules

No manual document creation is required: the app initializes `subjects`, `behaviors`, and `videos` as empty arrays and writes the first `updatedAt` value when the shared document is absent. To seed it manually, create a document with ID `shared` in the `lab_school_state` collection, containing those three fields as arrays and `updatedAt` as a string.

This app currently has no user authentication, so rules that permit its anonymous access necessarily make the shared data publicly readable and writable. For the single document, the minimum functional rules are:

```text
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /lab_school_state/shared {
      allow get: if true;
      allow create, update: if request.resource.data.keys().hasAll([
          'subjects', 'behaviors', 'videos', 'updatedAt'
        ])
        && request.resource.data.keys().hasOnly([
          'subjects', 'behaviors', 'videos', 'updatedAt'
        ])
        && request.resource.data.subjects is list
        && request.resource.data.behaviors is list
        && request.resource.data.videos is list
        && request.resource.data.updatedAt is string;
    }
  }
}
```

These rules intentionally do not permit listing the collection or deleting the document. Because any internet user can still read and change the shared document, use Firebase App Check and/or add authentication and stricter authorization before storing sensitive data.
If you set `VITE_FIREBASE_COLLECTION` to a non-default value, change the collection name in the rules to match.
