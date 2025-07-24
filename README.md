# Timing365 Outlook Add-in

This repository contains a React-based Office Add-in for Outlook. The project was
bootstrapped from the official Office Add-in TaskPane React template. It uses
React 17 and Fluent UI components.

## Project Structure

- `src/` – React source code
- `public/` – Office manifest and static assets
- `build/` – Webpack build output

## Scripts

- `npm run start` – Launch a development server at `https://localhost:3000`
- `npm run build` – Build the production bundle into the `build/` folder

The manifest at `public/manifest.xml` references `taskpane.html`, which is
created by the Webpack build.

## Configuring the manifest

`public/manifest.xml` acts as a template targeting the Office.js requirement set
1.13. Replace the following placeholders with values for your deployment:

- `{APP_ID}` – Unique add-in identifier (GUID)
- `{APP_NAME}` – Display name shown to users
- `{APP_DESC}` – Short description
- `{ICON_URL}` – Absolute URL to a 32×32 icon
- `{SUPPORT_URL}` – Link to a support page
- `{BUNDLE_URL}` – URL to the built `taskpane.html`

After running `npm run build`, host the generated bundle and update the
placeholders before sideloading the add-in.
