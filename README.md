# Column Views — Excel Add-in

A task-pane Excel Add-in that lets users create **named column-visibility views** for any worksheet. Switch between views with one click to instantly show or hide sets of columns — without losing the configuration.

Works on **Excel Desktop** (Windows & Mac) and **Excel Online**.

---

## Features

| Feature | Detail |
|---|---|
| Named views | Create as many views as needed per sheet |
| Per-sheet storage | Views are scoped to each worksheet |
| Persistent storage | Saved in the workbook's Custom XML Parts — survives save/close/reopen |
| Header detection | Automatically reads row-1 headers as column names |
| Show All | One-click to un-hide all columns |
| Edit / Rename | Update an existing view without losing others |
| Compatible | Excel Desktop (Win/Mac) + Excel Online |

---

## Project Structure

```
excel-column-views-addin/
├── manifest.xml          ← Office Add-in manifest (sideload this)
├── taskpane.html         ← Task pane UI
├── taskpane.css          ← Styles
├── taskpane.js           ← All add-in logic
├── commands.html         ← Required function-file stub
├── package.json          ← Dev-server scripts
└── assets/
    ├── icon-16.png       ← Ribbon icons (add your own)
    ├── icon-32.png
    └── icon-80.png
```

---

## Local Development Setup

### Prerequisites
- Node.js 16+
- A trusted local SSL certificate (required for Office.js)

### 1 — Generate a self-signed certificate

```bash
npx office-addin-dev-certs install
# Certificates are saved to ~/.office-addin-dev-certs/
```

Copy the generated `localhost.crt` and `localhost.key` into this folder, or point the start script at their actual location.

### 2 — Install dependencies

```bash
npm install
```

### 3 — Start the dev server

```bash
npm start
```

This serves the add-in at `https://localhost:3000`.

### 4 — Sideload in Excel

#### Excel Desktop — Windows
1. Open Excel → **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
2. Add `https://localhost:3000` as a trusted catalog **OR** use the Manifest sideload method:
   - Go to **Insert → Get Add-ins → My Add-ins → Upload My Add-in**
   - Select `manifest.xml`

#### Excel Desktop — Mac
1. Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel, then go to **Insert → My Add-ins**

#### Excel Online
1. Open a workbook in Excel Online
2. **Insert → Office Add-ins → Upload My Add-in**
3. Select `manifest.xml`

---

## Production Deployment

For production use, host the add-in files on any HTTPS server:

1. **Update `manifest.xml`** — replace all `https://localhost:3000` references with your production URL.

2. **Host options (recommended)**:
   - **Azure Static Web Apps** (free tier, built-in HTTPS)
   - **GitHub Pages** (free, HTTPS)
   - **Netlify / Vercel** (free tier, HTTPS)
   - Any static file host with HTTPS

3. **Distribute the manifest** via:
   - Microsoft 365 Admin Center (org-wide deployment) — *recommended for enterprises*
   - SharePoint App Catalog
   - Manual sideload for personal use

### Azure Static Web Apps (quickest)
```bash
# Install Azure CLI, then:
az staticwebapp create \
  --name column-views-addin \
  --resource-group my-rg \
  --source . \
  --location "eastus2"
```

---

## How It Works

### Storage
Views are stored in the workbook's **Custom XML Parts** — a hidden XML store that is part of the `.xlsx` file format. This means:
- Views travel with the workbook when shared
- No server or database required
- Works offline

### Column Detection
On opening the editor, the add-in reads the **used range** of the active sheet and detects:
- How many columns are in use
- Row 1 values (used as column display names)
- Current hidden/visible state (pre-fills checkboxes when creating a new view)

### Applying a View
When a view is applied, the add-in iterates over each saved column state and sets `columnHidden = true/false` via the Office JS API.

---

## Icons

Add 16×16, 32×32, and 80×80 PNG icons to the `assets/` folder. You can generate them from this SVG concept:

```svg
<svg viewBox="0 0 80 80" xmlns="http://www.w3.org/2000/svg">
  <rect x="5"  y="10" width="14" height="60" rx="4" fill="#4f7cff"/>
  <rect x="24" y="10" width="14" height="60" rx="4" fill="#4f7cff" opacity=".5"/>
  <rect x="43" y="10" width="14" height="60" rx="4" fill="#4f7cff"/>
  <rect x="62" y="10" width="13" height="60" rx="4" fill="#4f7cff" opacity=".3"/>
</svg>
```

---

## Extending the Add-in

### Row views
The same approach works for rows — replace `columnHidden` with `rowHidden` in `taskpane.js`.

### Keyboard shortcut
Add a `ExecuteFunction` control in `manifest.xml` bound to a keyboard shortcut to cycle views.

### Sheet Views integration
Excel's built-in **Sheet Views** (`workbook.worksheets[n].namedSheetViews`) can be used to layer column-hiding on top of a proper sheet view. This gives each user their own view without affecting others in collaborative sessions. To integrate:
```js
// Create a named sheet view, then apply column hiding within it
const view = sheet.namedSheetViews.add("My View");
await ctx.sync();
view.activate();
// ...apply column hiding...
```

---

## Browser / Platform Compatibility

| Platform | Supported |
|---|---|
| Excel Desktop Windows | ✅ |
| Excel Desktop Mac | ✅ |
| Excel Online (Browser) | ✅ |
| Excel Mobile (iOS/Android) | ⚠️ Task pane add-ins have limited support |

---

## License

MIT — use freely in personal and commercial projects.
