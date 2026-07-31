# Sheet Importer — Excel Office Add-in

A Microsoft Excel task-pane add-in that lets you import selected worksheets from external Excel workbooks directly into your active workbook — with a clean Fluent-UI interface.

---

## Features

- **Drag & drop** or browse to upload one or more `.xlsx` / `.xlsm` workbooks
- Automatically reads every worksheet in each uploaded file
- All sheets are **selected by default** — uncheck the ones you don't need
- Per-workbook cards with Expand/Collapse, Select All, Deselect All, and Remove
- Live summary: workbook count and selected sheet count
- **Progress bar** during import
- Sheet-name collision handling: `Sales → Sales (2) → Sales (3)`
- High-fidelity copy (charts, images, tables, conditional formatting) on Excel Desktop 2021 / Microsoft 365; data + formatting fallback on older builds

---

## Tech Stack

| Layer | Technology |
|---|---|
| Runtime | Office.js (Microsoft) |
| File parsing | SheetJS (xlsx.js) — CDN |
| UI components | Fluent UI Web Components — CDN |
| Language | HTML · CSS · JavaScript (ES6 modules) |
| Hosting | GitHub Pages |

No build step required. No Node.js. No React. No TypeScript.

---

## Project Structure

```
pull_sheets/
├── manifest.xml          # Office Add-in manifest (points to GitHub Pages)
├── index.html            # Redirect shim (required for manifest Commands URL)
├── taskpane.html         # Main add-in UI
├── taskpane.css          # Fluent-style stylesheet
├── taskpane.js           # Entry point — Office.onReady + event wiring
├── assets/
│   ├── js/
│   │   ├── workbook-manager.js   # State store (uploaded workbooks & sheets)
│   │   ├── file-reader.js        # SheetJS wrapper (read sheet names & data)
│   │   ├── importer.js           # Office.js import logic (high-fi + fallback)
│   │   └── ui.js                 # DOM rendering & updates
│   └── images/
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-80.png
└── README.md
```

---

## Local Development

### Prerequisites

- Any static file server (examples below)
- Excel Desktop (Microsoft 365 recommended) **or** Excel on the Web

### 1 — Clone the repo

```bash
git clone https://github.com/Luyandantombela/pull_sheets.git
cd pull_sheets
```

### 2 — Serve files locally with HTTPS

Office Add-ins require **HTTPS** even in development. The easiest options:

#### Option A — [live-server](https://www.npmjs.com/package/live-server) (simplest)

```bash
npx live-server --https
```

#### Option B — [http-server](https://www.npmjs.com/package/http-server) with a self-signed cert

```bash
# Generate a self-signed cert once
openssl req -newkey rsa:2048 -new -nodes -x509 -days 3650 \
  -keyout key.pem -out cert.pem

npx http-server -S -C cert.pem -K key.pem -p 3000
```

#### Option C — VS Code [Live Server extension](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer)

Enable HTTPS in the extension settings.

### 3 — Create a local manifest

Copy `manifest.xml` and replace all `https://luyandantombela.github.io/pull_sheets/` URLs with your local server URL (e.g. `https://localhost:3000/`).

```xml
<!-- Replace this: -->
<SourceLocation DefaultValue="https://luyandantombela.github.io/pull_sheets/taskpane.html" />

<!-- With this: -->
<SourceLocation DefaultValue="https://localhost:3000/taskpane.html" />
```

### 4 — Sideload the add-in into Excel

#### Excel Desktop (Windows)
1. Open Excel → **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
2. Add the path to the folder containing your local `manifest.xml`
3. Restart Excel → **Insert → My Add-ins → Shared Folder** → pick *Sheet Importer*

#### Excel Desktop (Mac)
1. Copy your local `manifest.xml` into `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel → **Insert → My Add-ins** → pick *Sheet Importer*

#### Excel on the Web (Office 365)
1. Go to **Insert → Add-ins → Upload My Add-in**
2. Browse to your local `manifest.xml` and upload

---

## Deploying to GitHub Pages (Production)

The `manifest.xml` already points to `https://luyandantombela.github.io/pull_sheets/`.

1. Push your code to the `main` branch of `https://github.com/Luyandantombela/pull_sheets`
2. In the repo **Settings → Pages**, set Source to **Deploy from a branch → main / root**
3. Wait ~1 minute for Pages to build
4. Visit `https://luyandantombela.github.io/pull_sheets/taskpane.html` to confirm it loads
5. Sideload `manifest.xml` (the unchanged production version) into Excel

---

## Import Fidelity

| Feature | High-fi path (M365 Desktop) | Fallback path |
|---|---|---|
| Cell values | ✅ | ✅ |
| Formulas | ✅ | ✅ |
| Number formats | ✅ | ✅ |
| Column widths | ✅ | ✅ |
| Row heights | ✅ | ✅ |
| Merged cells | ✅ | ✅ |
| Tables | ✅ | ❌ |
| Charts | ✅ | ❌ |
| Images | ✅ | ❌ |
| Conditional formatting | ✅ | ❌ |
| Named ranges | ✅ | ❌ |
| Cell styles / colours | ✅ | ❌ |

The high-fidelity path uses `insertWorksheetsFromBase64` (Excel Desktop 2021 / M365 only). The fallback is used automatically on older or web versions.

---

## Contributing

Pull requests are welcome. Please open an issue first for major changes.

## License

MIT
