# ANC Character Count Report Generator

Indha tool ungaloda `.db3` (SQLite) batch files padichu, screenshot la kaatiya
maadhiri "DocType" billing report (.xlsx) automatic ah build pannum.

## ⚠️ Important — EXE build panna Windows la than pannanum

Naan (Claude) இந்த மாதிரி work panra environment **Linux sandbox**. Adhula
irundhu neenga kekkura **Windows .exe** build panna mudiyadhu (cross-build
support pannadhu PyInstaller). So indha 3 files ah unga Windows PC ku kondu
poidunga, adhula `build_exe.bat` ah double-click pannunga — adhu automatic ah
`.exe` create pannidum.

## Files in this folder

| File | Purpose |
|---|---|
| `report_core.py` | Core logic — reads .db3 files, builds the .xlsx report |
| `app.py` | Tkinter GUI (folder picker + Generate button) |
| `requirements.txt` | Python packages needed (openpyxl, pyinstaller) |
| `build_exe.bat` | One-click builder — run this on Windows |

## How to build the .exe (on Windows)

1. Install Python 3.10+ (https://www.python.org/downloads/) if not already
   installed. While installing, tick **"Add Python to PATH"**.
2. Copy `report_core.py`, `app.py`, `requirements.txt`, and `build_exe.bat`
   into the **same folder**.
3. Double-click `build_exe.bat`.
4. Wait for it to finish — the exe will appear at `dist\DocTypeBillingReport.exe`.
5. You can copy just that one `.exe` file anywhere and run it — no Python
   needed after that (PyInstaller bundles everything).

## How to use the tool

1. Open `DocTypeBillingReport.exe`.
2. Click **Browse...** and select the folder that directly contains the
   `.db3` files (e.g. the extracted `62884.IDX.001` folder).
3. **Project ID** auto-fills from the folder name (e.g. `62884.IDX.001`) —
   you can edit it if needed.
4. Click **Generate Report...**, choose where to save, and you're done.

## What's inside the generated .xlsx

- **DocType** (1st sheet) — the detail report grid:
  - `Project ID`, `Doc Type Name`, `Batch Name`, `Image Count`,
    `Header Count`, `Record Count` — computed directly from counting rows in
    each batch's `Image` / `Header` / `Record` tables for that DocType.
  - `Character Count` — a **live Excel formula**:
    `=SUMPRODUCT(LEN(TRIM(RawData!<range>)))` over that batch+doctype's
    block of raw `_orig` field values, exactly as requested. If you ever
    hand-edit a value on the RawData sheet, this column recalculates
    automatically.
- **Summary** (2nd sheet) — one row per DocType, totalled across the whole
  project (all batches combined). `Image Count` / `Header Count` /
  `Record Count` / `Character Count` are all **live `SUMIF` formulas**
  pulling from the DocType sheet, so editing any value on DocType updates
  Summary automatically.
- **RawData** (hidden support sheet) — holds every `*_orig` field value per
  record, grouped in blocks per Batch + DocType. This is what the
  `SUMPRODUCT` formula on the DocType sheet points to. Unhide it if you ever
  need to double-check a formula's source range (right-click any sheet tab →
  Unhide).

## GUI

The desktop app is organized as tabs:
- **Tab 1 — "Manual Key"**: the DocType Billing Report generator described
  above (folder picker, Project ID, Generate button). Fully built.
- **Tab 2 — "Keying Partially Linked"**: placeholder for a future report
  type — shows "Coming soon" for now.

## Logic verified against your sample data (62884.IDX.001)

Naan indha zip file ah manually process pannu, ungaloda screenshot rows oda
ellam (content, folder, Probate — Image/Header/Record/Character counts)
**exact match** aachu. Example:
- Batch `...0024`, DocType `Content` → 302 / 302 / 367 / **4159** ✅
- Batch `...0022`, DocType `Folder` → 25 / 25 / 25 / **502** ✅
- Batch `...0006`, DocType `Probate` → 355 / 355 / 1307 / **31881** ✅

## Notes / assumptions

- Only DocTypes flagged `Indexed = 1` in each `.db3`'s `DocType` table are
  included in the report (Capture, Target, Bad, Duplicate, etc. are
  excluded — same as your sample).
- Row order: sorted by Batch Name, and within a batch, Folder → Content →
  other case types. Unga original sheet oda order konjam vera maari irukalam
  (adhu was processed in a different sequence originally) — but every
  value itself is verified correct.
- Doc Type Name is written exactly as stored in the database
  (`Content`, `Folder`, `Probate`, etc.) — unga screenshot la konjam mixed
  case irundhuchu (content/folder lowercase, Probate capital); if neenga
  specific casing venum na sollunga, adjust pannuven.

## 🔄 Auto-Update Setup ("Check for Updates" button)

Indha app la ippo "🔄 Check for Updates" button irukku (top-right, header la).
Adhu **GitHub Releases** ah check pannum — mobile app store update maadhiri:
button click pannina, latest version irukka nu check pannum, irundha
automatic ah download panni, current .exe ah replace panni, restart aagidum.

### Step 1 — GitHub la oru repo create pannunga (once only)

1. https://github.com ku poi login pannunga (account illana free ah signup
   pannunga).
2. Top-right "+" icon → **New repository**.
3. Repository name kudunga (e.g. `sps-tdm-billing-tool`).
4. **Public** ah vachunga (Private vachanaalum works, aana andha case la
   `updater.py` la ஒரு GitHub token add pannanum — ithu simple ah irukanumna
   Public recommend pandren, since idhu sensitive data illa, exe mattum than).
5. "Create repository" click pannunga.

### Step 2 — Unga local folder ah adha repo kooda link pannunga (once only)

Unga PC la, indha files ellam irukura folder ku poi (report_core.py, app.py,
updater.py, requirements.txt, build_exe.bat, icons, etc.), Command Prompt
open panni idha run pannunga:

```
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO>.git
git push -u origin main
```

(`git` install aagala na https://git-scm.com/download/win la irundhu download
pannunga first.)

### Step 3 — `updater.py` la unga repo peru set pannunga (once only)

`updater.py` file open panni, indha line ah edit pannunga:

```python
GITHUB_REPO = "YOUR_GITHUB_USERNAME/YOUR_REPO_NAME"
```

adha unga real username/repo ah maathunga, e.g.:

```python
GITHUB_REPO = "swiftprosys/sps-tdm-billing-tool"
```

### Step 4 — Ovvoru puthu version varum bodhum idha pannunga (repeat every release)

1. `app.py` la `APP_VERSION = "1.0.0"` ah next version ku maathunga
   (e.g. `"1.1.0"`).
2. `build_exe.bat` double-click panni puthu `.exe` build pannunga
   (`dist\SPS_TDM_DocTypeBillingReport.exe` varum).
3. GitHub repo page ku poi → right side **"Releases"** → **"Create a new
   release"**.
4. **Tag**: `v1.1.0` (APP_VERSION oda same number, front la `v` podunga).
5. **Title**: e.g. "Version 1.1.0".
6. **Description**: enna changes pannirukinga nu type pannunga (idhu than
   app la "New version available" popup la kaamikkum).
7. **Attach binaries**: dist folder la irundha `.exe` file ah drag & drop
   pannunga.
8. **"Publish release"** click pannunga.

Idhudhaan! Ippo yaaru vendumaanaalum antha app open panni "🔄 Check for
Updates" click pannina, puthu version irukka nu automatic ah therinjidum,
download panni, replace panni, restart aayidum — mobile app update maadhiri.

### Idhu epidi internally work aagudhu

- Button click pannina, `updater.py` GitHub oda "latest release" API ah check
  pannum (`https://api.github.com/repos/<repo>/releases/latest`).
- Adhula irukura tag (e.g. `v1.1.0`) ah, app la irukura `APP_VERSION` kooda
  compare pannum.
- Puthusa irundha, andha release oda attached `.exe` ah download pannum.
- Windows la, currently running `.exe` ah direct ah replace panna mudiyadhu
  (file lock irukkum), so oru chinna `.bat` script create panni, adhu:
  current app close aagi wait pannitu → old exe delete pannitu → puthu exe
  antha peru la (same filename) rename pannitu → restart pannidum. App ippo
  `os._exit(0)` panni immediate ah close aagidum, adhanaala andha `.bat`
  step successful ah nadakkum.
- **Note**: idhu compiled `.exe` la mattum than work aagum — `python app.py`
  nu source la run pannina, "Self-update only works in the built .exe" nu
  message varum.
