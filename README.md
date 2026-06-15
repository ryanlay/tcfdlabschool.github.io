# Lab School Video Behavior Database React Clone

This is a React + Vite clone of the Lab School app, built to run as a static site on GitHub Pages.

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

The app uses `base: './'`, so the generated `dist/` folder can be hosted on `github.io` as a project site.

## Features

- Intake wizard with subjects and behavior selection
- Query-style review tables
- Searchable data view
- Admin lists for subjects and behaviors
- Persistent data in `localStorage`
- SharePoint folder links based on the recording date
