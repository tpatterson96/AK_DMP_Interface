[ReadMEv3.md](https://github.com/user-attachments/files/32308248/ReadMEv3.md)
# AK Region DMP Interface

This replaces `AK_DMP_Template_v1_3.docx` as the way projects fill out a Data
Management Plan. Instead of typing into Word content-control boxes, you fill
out a form in your browser (or in a small desktop app — see below).
Everything else about the pipeline — the contact-matching, the mdEditor JSON
generation, and the National DMP SharePoint export — is unchanged, and still
happens in `DMP_to_metadata_from_interface.ipynb`.

## What's in this folder

| File | Purpose |
|---|---|
| `DMP_Interface.html` | The form itself. Open it (or double-click the launcher) to fill out a DMP. |
| `Launch_DMP_Interface.bat` | Double-click to open the interface in your default browser. |
| `dmp_json_to_dataframe.py` | Bridge script. Turns the interface's JSON export into the same dataframe the notebook used to get by scraping a Word doc. |
| `DMP_to_metadata_from_interface.ipynb` | A copy of `DMP_to_metadata.ipynb` with the Word-scraping cells swapped for the JSON bridge. Everything from "Pulling contacts" onward is untouched. |
| `region_program_contacts.py` | Resolves the interface's "FWS Region" and "FWS Program" selections into mdEditor contact UUIDs. IDs are hardcoded (see below) — no CSV file is required. |
| `FWSRegion_Program_Contacts_mdeditor-*.json` | mdEditor export of the FWS Region and FWS Program organization contact records. Loading this on the Setup tab lets the *full* contact record be embedded in generated metadata, not just a bare ID. |
| `repository_options.csv` | Editable list of "Public Data Sharing Repositories" options for the dropdown in Preservation & Distribution. Edit this in Excel and load it into the interface's Setup tab to update the list. |
| `dmp_desktop_app.py` | Optional desktop wrapper (pywebview). Runs the same interface in a native window with real Python behind it — see "Desktop app" below. |
| `requirements.txt` | Python packages needed for the desktop app (`pywebview`, `pandas`). |
| `Setup_And_Run_Desktop_App.bat` | Double-click to install the desktop app's requirements (first run only) and launch it. |

**Not shipped in this folder, but referenced by the interface:** `CMT.csv`
(your organization's cost center directory — large and specific to your
region, so it isn't checked into this repo). You load it once on the Setup
tab; see "Cost Center filtering" below.

Nothing here requires a server or an internet connection to run. The `.html`
file is self-contained — CSS and JavaScript are inline, there are no CDN
calls or external fonts — so opening it via `file://` works the same on a
locked-down federal machine as anywhere else.

## Filling out a DMP

1. Double-click `Launch_DMP_Interface.bat` (or open `DMP_Interface.html`
   directly in Chrome/Edge/Firefox) — or use the desktop app instead (see
   below).
2. Work through the eleven numbered sections in the left sidebar — they
   follow the same order and required fields (marked with `*`) as the old
   Word template, plus two additions:
   - **Purpose** (Project Details, optional) — a brief statement of *why*
     the project or its data are being collected, distinct from the
     Abstract (which summarizes *what* the project is). Maps to the mdJSON
     `metadata.resourceInfo.purpose` field.
   - **Preservation & Distribution** and **Records Schedule** are now two
     separate tabs (Sections 6 and 7) instead of one combined tab.

   Repeatable items (extra contacts, products, originators, records
   schedules, etc.) have a **"+ Add"** button, matching the old template's
   `[+]` content-control behavior.
3. A **⚙ Setup** entry sits below the numbered sections, visually separated
   as "Configuration" — this is where you point the interface at your
   organization's shared files (see "Setup tab" below). Nothing entered
   there is written into your DMP data; it's pipeline configuration, kept
   separate on purpose.
4. Your work autosaves to the browser's local storage as you type, so you
   can close the tab and come back later on the same machine. Use
   **"Load JSON…"** to resume a draft you exported earlier, or to hand a
   partially-completed plan to a colleague.
5. On the **Review & Export** screen, check the completeness list, then use
   one of the export options (see "Exporting" below).

## Setup tab

Everything the interface can load from an external file lives here, in one
place, so it isn't scattered across sections:

- **Public Repository Dropdown List** — the same `repository_options.csv`
  mechanism as before.
- **Cost Center Directory (`CMT.csv`)** — see "Cost Center filtering" below.
- **Contacts Export** — an mdEditor contacts export, used to match this
  DMP's personnel by email during draft metadata generation (see
  "Exporting" below) so real contact UUIDs get reused instead of
  duplicated.
- **FWS Region / Program Contacts** — the full mdEditor contact *records*
  for FWS Regions/Programs (not just an ID — see next section). Optional:
  role references still work without it, using just the hardcoded ID.
- **Profile & Schema Attachments** — optional; if enabled, fetches the
  mdEditor profile/schema JSON from GitHub and includes it in a draft
  metadata export for reference. Requires outbound internet access.
- **Python Pipeline Paths** — reference locations (bridge script folder,
  contacts directory, template folder, spreadsheet path) for the full
  notebook pipeline. In the browser build these are copy-paste text only —
  the browser can't run Python. In the desktop app, these are used
  directly (with native **Browse…** buttons) and a **Copy to clipboard**
  button generates a ready-to-paste settings block for the notebook.

All Setup tab choices persist in the browser's local storage on that
computer, separately from `dmp_data.json`, so they survive "Clear form" and
don't need to be re-entered every session.

## FWS Region and FWS Program contacts

Every metadata record generated by the notebook (or by the in-browser/desktop
draft generators — see "Exporting" below) lists the project's **FWS Region**
as an `administrator`, `distributor`, and `publisher` contact, and each
selected **FWS Program** as an `administrator` contact.

This is resolved by `region_program_contacts.py` using **hardcoded mdEditor
contact IDs** (`REGION_CONTACT_ID` / `PROGRAM_CONTACT_ID` near the top of
that file) — sourced from a real mdEditor contacts export, not placeholders.
No CSV file is required for this to work. A name-based fallback
(`REGION_CONTACT_NAME` / `PROGRAM_CONTACT_NAME`) is tried only if a
hardcoded ID isn't found in whatever contacts export is currently loaded.

If your organization's mdEditor contact record for a Region or Program is
ever recreated (getting a new contactId), or a new Region/Program is added,
update the relevant table in `region_program_contacts.py` to match — and
mirror the same change in `DMP_Interface.html`'s `REGION_CONTACT_ID` /
`PROGRAM_CONTACT_ID` JavaScript objects, so the browser and desktop-app
draft generators stay in sync with the notebook.

If a selected region or program doesn't resolve by either method, the
notebook prints a warning and simply skips adding that one — it won't stop
the rest of the notebook from running.

**Headquarters:** the FWS Region dropdown includes a 9th option, "HQ –
Headquarters," in addition to Regions 1–8. It resolves to the real
Headquarters mdEditor contact, and — since `CMT.csv` uses `9` in its
`Region` column for Headquarters rows — it filters Cost Centers the same
way any other region does (see below).

## Cost Center filtering

Project Details' Cost Center dropdown has **no built-in list** — it stays
empty until you load `CMT.csv` on the Setup tab. This is intentional:
fabricating cost center codes would be worse than requiring one real load
per session.

Once `CMT.csv` is loaded:
- The dropdown filters to rows matching the selected **FWS Region** (via
  `CMT.csv`'s `Region` column; Regions 1–8 map directly, Headquarters maps
  to `9`).
- It further narrows by **FWS Program**, using each row's `CMT_CCCode5`
  column. This mapping is **derived automatically** — the interface scans
  `CMT.csv`'s own `Program` and `CMT_CCCode5` columns and matches them
  against the six FWS Program dropdown strings, picking the most common
  letter for each program. No separate mapping file is needed.
- The **derived mapping is shown on the Setup tab** (Program / Derived
  letter / Matched row count) so you can visually verify it before relying
  on it. If a program comes up unmatched (e.g. `CMT.csv`'s wording for that
  program differs from the dropdown), you'll see a warning there — you can
  force a specific mapping by adding an entry to
  `PROGRAM_CMT_CODE_OVERRIDE` near the top of `DMP_Interface.html`'s
  script.
- Rows explicitly marked inactive in `CMT.csv`'s `Active` column are
  excluded.

`CMT.csv` needs at minimum: `CCCode`, `CMT_CCCode5`, `Region`, `OrgName`,
`Program`, and `Active` columns.

## Updating the repository dropdown list

The "Public Data Sharing Repositories" dropdown (Preservation &
Distribution tab) is driven by `repository_options.csv` instead of being
hardcoded. To update it:

1. Open `repository_options.csv` in Excel (or any spreadsheet app). Column A
   is the repository name shown in the dropdown; column B is an optional
   description for whoever maintains the file next — it isn't shown in the
   app. Keep the header row; it's always skipped when the file is read.
2. Add, remove, or rename rows as needed, then save.
3. On the **Setup** tab, in the "Public Repository Dropdown List" card,
   click **Load CSV…**, then select the updated file.

That updates the dropdown immediately and remembers your choice (via the
browser's local storage) the next time you open the interface on that same
computer. Use **Reset to default** to go back to the built-in list at any
time.

**A note on how this loads:** because the interface is a single offline
file with no server, it can't silently read a CSV off disk on its own when
opened via double-click (`file://`) — browsers block that for security
reasons, so **Load CSV…** / **Browse file…** is the reliable way to update
any of the Setup tab's files. If your organization ever hosts this folder
on an internal website instead of distributing it as a local file, the
interface will automatically pick up the latest `repository_options.csv`
on every page load with no manual step needed (this auto-fetch is specific
to that one file for historical reasons; everything else on Setup uses the
manual loader).

## Exporting

The Review & Export tab offers several export options, in increasing order
of how "finished" the output is. Each produces a file named from your
**Project Title** (Section 2), sanitized to letters, numbers, and
underscores — if no title has been entered yet, `untitled` is used in its
place.

| Button | Filename pattern | Example (title "Walrus Haulout Monitoring") |
|---|---|---|
| Export DMP | `DMP_PDF_<title>.pdf` *(suggested — see note below)* | `DMP_PDF_Walrus_Haulout_Monitoring.pdf` |
| Export draft metadata (browser build) | `DMP_draft_<title>.json` | `DMP_draft_Walrus_Haulout_Monitoring.json` |
| Run metadata draft creator (desktop app) | `DMP_draft_<title>.json` | `DMP_draft_Walrus_Haulout_Monitoring.json` |
| Export data file (.json) | `DMP_JSON_<title>.json` | `DMP_JSON_Walrus_Haulout_Monitoring.json` |

- **Export DMP** — opens your browser's print dialog with a formatted,
  report-style rendering of the whole plan. Choose "Save as PDF" as the
  destination to get an actual PDF — no library needed, still fully
  offline-safe. This is the closest equivalent to printing the old Word
  document. The interface sets the page title to the suggested filename
  before printing, so most browsers (Chrome/Edge) will pre-fill it in the
  "Save as PDF" dialog — but unlike the JSON exports, this isn't
  guaranteed, since the print dialog is controlled by the browser, not the
  page.
- **Export draft metadata** (browser build) — runs a JavaScript port of
  `dmp_json_to_dataframe.py`'s field mapping and produces real mdJSON
  content (project + product records, resolved contacts, Region/Program
  admin contacts), useful for spot-checking before running the real
  pipeline. It does **not** reproduce mdEditor's own internal
  record/profile/schema envelope, extract a PRIMR ID, embed geoJSON
  geometry, or write the SharePoint Excel export — those still require the
  notebook.
- **Run metadata draft creator** (desktop app only) — calls the *actual*
  `dmp_json_to_dataframe.py` and `region_program_contacts.py` directly (no
  JavaScript approximation), and saves the result via a native file dialog.
  Still a draft in the same sense as above (no mdEditor envelope, no
  SharePoint export) — but the field mapping and contact resolution are
  running the real Python, not a port of it. Uses the same filename
  pattern as the browser build's draft export, since both produce the same
  kind of file.
- **Export data file (.json)** — the file that actually matters for the
  full pipeline. Hand this to the notebook (see below) for the fully
  resolved, import-ready mdEditor file and the SharePoint list export.

Renaming the downloaded file afterward is always fine — nothing about the
filename itself is read by the notebook or any other tool; `dmp_json_path`
in the notebook settings just needs to point at wherever you saved it.

## Feeding the exported data file into the metadata pipeline

You have two options:

**Option A — use the pre-modified notebook (recommended).**
Open `DMP_to_metadata_from_interface.ipynb` instead of the original
notebook. In the first cell, set:

```python
dmp_json_path = r'C:\path\to\your_exported_file.json'
bridge_dir    = r'C:\path\to\this_folder'   # wherever dmp_json_to_dataframe.py lives
region_program_contacts_path = r'C:\path\to\this_folder\FWSRegion_Program_Contacts_mdeditor-....json'
```

along with the same `contact_folder`, `tmp_dir`, `dmp_spreadsheet`, and
profile/schema settings you were already using. (The Setup tab's "Python
Pipeline Paths" card can generate this block for you — fill in the fields
there and click **Copy to clipboard**.) Run the notebook top to bottom as
before — the contacts matching, mdEditor JSON generation, and SharePoint
export cells are byte-for-byte identical to the original.

**Option B — patch your existing notebook.**
If you'd rather keep using `DMP_to_metadata.ipynb` directly, replace the
cells that open `blankdmp`/`dmp` and scrape `word/document.xml` with:

```python
import sys
sys.path.append(r'C:\path\to\this_folder')
import dmp_json_to_dataframe as bridge

df, dmp, dmpvers = bridge.build_dataframe(r'C:\path\to\your_exported_file.json')
```

Everything from the "Pulling contacts" section onward reads `df`, `dmp`, and
`dmpvers` exactly the same way it always did, so no other cells need to
change.

## Desktop app (optional)

If Python is available on the machine, you can run the interface as a
small native desktop app instead of a browser tab — this lets buttons in
the interface actually invoke real Python (something a plain browser page
structurally cannot do, regardless of settings).

**First-time setup:** put `dmp_desktop_app.py`, `requirements.txt`, and
`Setup_And_Run_Desktop_App.bat` in this same folder, then double-click the
`.bat` file. It installs `pywebview` and `pandas` (skipped automatically on
future runs) and opens the app in a native window.

**What's different in the desktop app:**
- The Setup tab's Python Pipeline Paths gain real **Browse…** buttons
  (native folder/file pickers) instead of plain text fields.
- Review & Export gains a **"Run metadata draft creator"** button that
  calls the real bridge scripts directly and saves via a native Save
  dialog (see "Exporting" above).
- Everything else — filling out the form, Cost Center filtering, the
  Setup tab's other cards — behaves identically to the browser build.

If Python isn't installed on the target machine, or your IT policy doesn't
allow installing packages, the plain browser build (`DMP_Interface.html`)
remains fully functional on its own — the desktop app is an enhancement,
not a requirement.

## Why this works

The old notebook's first several cells did one job: turn a filled-out Word
document into a pandas dataframe with columns `field`, `value`, `prod_num`,
`sample_num` — one row per content-control box, with product/sample/contact
entries numbered to keep repeated sections apart (e.g. a project with two
data products ends up with `productTitle1`, `productTitle2`, and two
originators on the second product become `productOriginatorFirstName2-1`,
`productOriginatorFirstName2-2`).

`dmp_json_to_dataframe.py` reproduces that exact same field-naming and
numbering scheme from the interface's JSON, field-for-field, so the
dataframe it produces is indistinguishable from what scraping a real `.docx`
used to produce. That's what lets every downstream cell — contact
de-duplication against your existing contacts file, mdEditor record
construction, the SharePoint spreadsheet export — run without modification.

## If a field is missing or wrong

If FWS adds a field to a future version of the DMP template, or you notice a
mismatch between what the interface exports and what the notebook expects,
the fix lives in one place: `dmp_json_to_dataframe.py`. Each section of that
file mirrors one section of the interface's JSON output, in the same order
the fields appeared in the original Word template — cross-reference against
`AK_DMP_Template_v1_3.docx`'s content-control tags if you need to add one.

The **Purpose** field (added under Project Details) is the most recent
field addition and a good template to follow if you need to add another: it
required a matching change in three places — `DMP_Interface.html` (the
form field, `bridgeBuildRows()`, and the draft mdJSON builder),
`dmp_json_to_dataframe.py` (`projectPurpose`), and `dmp_desktop_app.py`
(`_assemble_draft()`'s `resourceInfo` dict).

Two fields predate the original Word template and don't map to a
content-control tag: **FWS Region** (top of Project Details → Program &
Cost Center) and **Storage location URL** (top of Storage, Backup &
Review). Both export as columns (`fwsRegion`, `storageLocationURL`) in the
dataframe; neither is currently read by any downstream cell, so they're
informational unless you wire them into the mdEditor or SharePoint export
logic yourself.

If you notice an exported filename that doesn't match the patterns
described above in "Exporting," the fix lives in `DMP_Interface.html`: each
export button has its own small filename-construction line right at the
top of its handler function (`doExport()`, `doExportDraftMetadata()`,
`exportDmpDocument()`, and — for the desktop app — `dmp_desktop_app.py`'s
save handler). All four should stay in sync with each other and with this
README if the naming convention ever changes.
