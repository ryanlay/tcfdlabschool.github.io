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

- Intake wizard with subjects and behavior selection
- Query-style review tables
- Searchable data view
- Admin lists for subjects and behaviors
- Shared persistent data in SharePoint (Microsoft Graph)
- SharePoint folder links based on the recording date

## Shared Database Setup (SharePoint)

This app is set up to store `subjects`, `behaviors`, and `videos` in a shared SharePoint list.

Important: direct browser requests from GitHub Pages to SharePoint are blocked by CORS/auth rules in most tenants. For true cross-device sync, the app must either:

- be hosted on the SharePoint origin itself, or
- use a backend / Entra app registration / proxy that can talk to SharePoint securely.

If the app is served from GitHub Pages without that extra hosting/auth layer, it will not be able to reliably read/write the shared list.

### 1) Register an app in Microsoft Entra ID (recommended for GitHub Pages)

- Create an app registration.
- Add SPA redirect URIs:
  - `http://localhost:5173`
  - `https://ryanlay.github.io/tcfdlabschool.github.io/`
- Grant Microsoft Graph delegated permission: `Sites.ReadWrite.All`
- Grant admin consent for your tenant.

### 2) Create local environment file

Create `.env.local` in the project root:

```env
VITE_AAD_CLIENT_ID=your-app-client-id
VITE_AAD_TENANT_ID=your-tenant-id
VITE_SHAREPOINT_HOSTNAME=thecenterfordiscovery.sharepoint.com
VITE_SHAREPOINT_SITE_PATH=sites/LabSchool
VITE_SHAREPOINT_LIST_NAME=LabSchoolAppState
```

### 3) SharePoint list used for app state

The app auto-creates the list if it does not exist (named by `VITE_SHAREPOINT_LIST_NAME`) and stores one row titled `SharedState` containing JSON fields:

- `subjectsJson`
- `behaviorsJson`
- `videosJson`

## Notes

- GitHub Pages hosts static files only; cross-origin SharePoint writes are blocked unless you add an auth/proxy layer or host the app on SharePoint.
- Users must be signed in with permitted Microsoft accounts to read/write shared data.
- If you need, I can help switch this project to a supported backend path next.
