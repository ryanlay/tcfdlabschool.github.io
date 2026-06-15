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
- Persistent data in `localStorage`
- SharePoint folder links based on the recording date
