# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ระบบตรวจสอบอุปกรณ์ดับเพลิง — a fire safety equipment inspection system built on **Google Apps Script + Google Sheets**. The app runs as a GAS Web App (standalone URL) and also as a Sheets add-on menu.

## Deployment Commands

```bash
# Push local src/ to Google Apps Script
clasp push

# Pull latest from GAS editor back to local
clasp pull

# Open project in GAS online editor
clasp open

# Check push/pull status
clasp status
```

All source files are in `src/`. The `rootDir` in `.clasp.json` points there. Every file in `src/` is deployed — `.gs` files become server-side scripts, `.html` files become HtmlService templates.

There is **no build step, no bundler, no test runner**. Testing is done manually by running the web app or calling GAS functions from the editor console.

## Architecture

### Data layer — Google Sheets as DB

Each Sheet tab = one table. Column A is always `id` (UUID v4 via `Utilities.getUuid()`). All reads/writes go through `SheetHelper.gs`:

- `getAllAsObjects(sheetName)` — batch read entire sheet → array of plain objects
- `appendRowGetId(sheetName, rowData)` — insert row, returns new UUID
- `findRowById(sheet, id)` — locate 1-based row index by id
- `updateRowById / deleteRowById` — batch update/delete by id

Never call `SpreadsheetApp` directly from entity files — always go through SheetHelper.

### Entity files (CRUD layer)

Each entity has its own `.gs` file: `Building.gs`, `Floor.gs`, `Location.gs`, `Department.gs`, `Equipment.gs`, `Inspection.gs`. Each follows the same pattern: `addX`, `getX`, `getAllX`, `updateX`, `deleteX`. Delete functions enforce FK constraints (e.g., `deleteBuilding` checks for child floors).

FK chain: **Building → Floor → Location ← Inspection → Equipment, Department**

### Web App entry points (`Code.gs`)

- `doGet()` — serves `Index.html` as the web app
- `getDashboardData()` / `getMasterData()` / `getDropdownData()` — data fetch functions called from the frontend
- `web*` prefixed functions (e.g., `webAddBuilding`, `webDeleteFloor`) — thin CRUD wrappers that catch errors and return `{ success, error }` objects
- All config stored in `PropertiesService.getScriptProperties()` (not hardcoded)

### Frontend (`Index.html`)

Single-page app with four views: Dashboard, Inspection Form, Master Data, Settings. Navigation is client-side only (no page reloads). All server calls use the GAS `google.script.run` pattern:

```js
google.script.run
  .withSuccessHandler(fn)
  .withFailureHandler(fn)
  .serverFunction(args);
```

Output from server functions must be JSON-serializable (no Date objects — use formatted strings).

### Dashboard (`Dashboard.gs`)

`refreshDashboard()` writes aggregated data into the Dashboard sheet. `sendOverdueAlert()` sends Gmail alerts for overdue equipment. Both are called by `dailyJob()` via a time-driven trigger.

## Key Constants (`Code.gs`)

Sheet names and result enums are constants — always use them, never hardcode Thai strings:

```js
SHEET_INSPECTION, SHEET_BUILDING, ...   // sheet tab names
RESULT_PASS, RESULT_FAIL, RESULT_REPAIR // 'ผ่าน', 'ไม่ผ่าน', 'ซ่อมแซม'
DEPT_TYPE_OWNER, DEPT_TYPE_INSPECTOR    // department type enums
CONFIG_SPREADSHEET_ID, CONFIG_ALERT_EMAIL // PropertiesService keys
```

## HTML Inclusion Pattern

GAS does not support external CSS/JS files. To split `Index.html`, use:

```html
<?!= HtmlService.createHtmlOutputFromFile('Style').getContent(); ?>
<?!= HtmlService.createHtmlOutputFromFile('Script').getContent(); ?>
```

`Style.html` wraps content in `<style>` tags; `Script.html` wraps in `<script>` tags.

## Web App Config

- Runtime: V8, timezone: Asia/Bangkok
- Executes as: USER_DEPLOYING, access: ANYONE_ANONYMOUS
- OAuth scopes: spreadsheets, gmail.send, script.scriptapp, script.send_mail
