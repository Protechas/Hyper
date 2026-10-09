# Hyper

## HTML/CSS desktop interface

Start the redesigned interface on Windows with `powershell -ExecutionPolicy Bypass -File .\Start-Hyper.ps1`.
The launcher creates a local `.venv` and installs dependencies on first use. It supports Python 3.11
from python.org, the Windows `py` launcher, or the Microsoft Store aliases (`python.exe`,
`python3.11.exe`, and `python3.exe`). Each candidate is executed and version-checked rather than
trusted by path alone. If a copied `.venv` points to Python on another computer, the launcher
automatically rebuilds it using the first working Python 3.11 installation it finds.
Alternatively, install `requirements.txt` in your existing Python environment and run `python HyperWeb.py`.

The new front end has light and dark themes, responsive selection cards, keyboard accessible controls,
and the existing compact progress view. Its theme preference is stored locally in the embedded browser.
On Windows 11, the native title bar and window border follow the crimson palette; native window
controls and resize behavior remain available. Earlier Windows versions retain their native frame.
The header uses a custom linked research graphic. Brief entrance and hover highlights respect the
system's reduced-motion preference.
`ui/index.html`, `ui/styles.css`, and `ui/app.js` supply the presentation; `HyperWeb.py` connects it to
the original PyQt widgets through Qt WebChannel. Labels, defaults, checkbox signals, file dialogs,
sign-in validation and attempt limits, confirmation dialogs, automation workers, reports, and run
controls use the original code. Upload destination controls remain hidden as in the original GUI.

While automation is running, click **Collapse to icon**, minimize the window, or close it with the
title-bar X to keep the job running behind a small floating hyperlink icon. Drag the icon to move
it. Click it to restore the progress window. Right-click it for the existing Pause/Resume and Stop
controls; Stop still uses the original confirmation. A status bubble appears above the icon for
six seconds, about every 30 seconds, on a phase change (at most once per 15 seconds), and immediately
on pause/resume or completion. Hover over the icon for a current update. Closing Hyper while idle
still exits normally. The icon remains available after completion until the window is restored.

`Hyper.py`, `SharepointExtractor.py`, and `pdf_annotation_extractor.py` are unchanged. Running
`Hyper.py` still opens the original interface. Opening the HTML file directly cannot run automation;
use the desktop launcher. Native confirmation, file, and log dialogs retain the original Qt styling.

To verify the front end without starting SharePoint automation, run `python test_web_ui.py`.
Screenshots are written to `work/screenshots`, or the folder specified by `HYPER_SCREENSHOTS`.

Hyper is a Windows desktop automation tool used by Protech Automotive Solutions to connect OEM service-information PDFs stored in SharePoint with the correct rows in Excel long sheets. It supports ADAS Service Information, Repair Service Information, multiple manufacturers and year ranges, broken-link repair, and a separate PDF upload workflow.

Hyper does not read a PDF's technical contents to decide where its link belongs. The normal hyperlink workflow reads the SharePoint folder structure and PDF filename, extracts the year, manufacturer, model, and system acronym, and matches those values to workbook headers and rows.

> **Important:** Hyper edits selected Excel workbooks in place. Work from a backed-up or versioned copy until the results have been reviewed.

## Main components

| File | Responsibility |
| --- | --- |
| `Hyper.py` | PyQt5 login and main window, selections, run queue, progress, pause/stop controls, retry handling, logs, and reports. |
| `SharepointExtractor.py` | Chrome/SharePoint navigation, file discovery, link generation, filename parsing, workbook matching, hyperlink placement, cleanup mode, and uploading. |
| `pdf_annotation_extractor.py` | Upload-mode preprocessing: finds annotated pages, combines multipart documents, filters unwanted material, and creates a processing report. |
| `settings.json` | Current UI preference storage. |

`data.db` exists in the repository but is not referenced by the current Python workflow.

## Requirements and startup

- Windows and Python 3.11 are the environment used by this repository.
- Google Chrome must be installed.
- The Windows user must have access to the Caliber/Protech SharePoint locations.
- The user's Chrome profile must already be signed in to Microsoft 365/SharePoint.
- Excel workbooks must be `.xlsx` files following the conventions below.
- Python packages include PyQt5, Selenium, OpenPyXL, PyMuPDF, PyAutoGUI, pywin32, psutil, `chromedriver-autoinstaller`, and `webdriver-manager`.

The repository contains a virtual environment named `env`. On the configured workstation, start Hyper from PowerShell with:

```powershell
.\env\Scripts\python.exe .\Hyper.py
```

If that environment is not portable to another workstation, create a Python 3.11 virtual environment and install the dependencies above.

## Sign-in Verification

There are two separate sign-in layers:

1. **Hyper application sign-in.** The code currently validates the local username `Dromero221` (the Romero221 account) before opening the main window. It allows five attempts. This check is local to `Hyper.py`; it is not Microsoft authentication.
2. **Microsoft 365/SharePoint sign-in.** Selenium uses a copied Chrome profile to reuse an existing SharePoint session. Hyper first looks for Chrome's `Default` profile and otherwise tries `Profile 1`. On the first run it copies that profile into `%USERPROFILE%\ChromeAutomationProfiles` and uses the copy for automation.

The local Hyper password is currently embedded in source code and is intentionally not duplicated here. For a shared or production deployment, move both local credentials into environment variables or an approved authentication/secret system.

Because the automation profile is copied only when it does not already exist, a stale Microsoft session may require an administrator to refresh the copied profile. Never remove a profile while Hyper or Chrome is using it.

## Normal hyperlink workflow

1. Hyper collects the selected Excel files, manufacturers, year ranges, and systems.
2. The ADAS/Repair switch selects the corresponding system list and SharePoint roots.
3. The confirmation dialog shows exactly what will run.
4. Hyper pairs selected manufacturers and Excel files by position.
5. For each manufacturer, Hyper processes the selected year-range roots one at a time.
6. `SharepointExtractor.py` opens SharePoint in Chrome, enters manufacturer folders, scrolls virtualized lists, and recursively indexes folders and files.
7. It filters files to the selected systems and generates a SharePoint sharing link for each accepted document.
8. It opens the workbook with OpenPyXL, detects columns from row-one headers, parses each filename, locates the best row, and places the link.
9. The workbook is saved in place, and Hyper advances to the next range and manufacturer.
10. Failed manufacturers are retried later, up to ten attempts, and a final report is written.

### File/manufacturer ordering is significant

Selected Excel files and checked manufacturers are parallel lists. The first workbook is used for the first selected manufacturer, the second workbook for the second manufacturer, and so on. Keep both lists in the same order and provide one workbook per manufacturer. If there are fewer files than manufacturers, the unmatched manufacturer is skipped.

### Supported year ranges

The UI presents 2012–2016, 2017–2021, 2022–2026, and 2027–2031. The first three have normal ADAS and Repair SharePoint source links configured. The normal source lists do not currently contain a fourth link, and the 2027–2031 upload destinations are placeholders/incomplete. Do not expect that range to work until real URLs are configured.

## How filenames are read

The legacy naming pattern is:

```text
YYYY Manufacturer Model (SYSTEM).pdf
```

The new ADAS naming pattern is:

```text
YYYY Manufacturer Model (SYSTEM) Calibration Type [Component Name].pdf
```

Examples:

```text
2024 Honda Accord (ACC 1).pdf
2018 Chrysler Pacifica [PHEV] (BUC) No Cal Req [Backup Camera].pdf
2018 Chrysler Pacifica [PHEV] (SVC) Dynamic [Surround View Camera].pdf
2024 Ford F-150 (WSR) Static or Dynamic [Windshield Radar/LIDAR].pdf
2022 Kia Niro EV (SAS).pdf
2021 BMW X5 (Steering Angle Sensor).pdf
```

Hyper parses a filename in this order:

1. The first `20xx` value becomes the model year.
2. The known manufacturer name is removed.
3. Remaining words before a recognized system token become the model.
4. A recognized acronym in parentheses is preferred as the system.
5. If no recognized parenthesized acronym exists, Hyper searches the whole filename, then square brackets, then the final token.
6. In the new format, text after the system is parsed as the calibration type and the final square-bracket value is treated as the component description, not part of the model.
7. If the model cannot be recovered, Hyper can use the parent folder name.

Parentheses containing model qualifiers such as `(HEV)`, `(PHEV)`, or `(EV)` remain part of model matching unless recognized as a system. Square-bracket qualifiers such as `[HEV]` are ignored during relaxed comparison. System punctuation and spacing are normalized, so `ACC 2`, `ACC-2`, and `ACC2` resolve to `ACC2`.

Recognized calibration values are `Static`, `Dynamic`, `Static or Dynamic`, `Static & Dynamic`, `Program`, `Initial`, `Verify`, and `No Cal Req`. Matching is case-insensitive and the saved workbook value uses this standardized spelling.

Consistent filenames produce the safest placements. A missing year or unrecognized system prevents an exact match and may produce a review-required fallback.

## ADAS systems and acronym conversion

The **Old/New** toggle changes both the visible ADAS checklist and the workbook column used for placement. Hyper starts in **Old** mode.

- **Old** shows `ACC`, `AEB`, `FCW`, `AHL`, `APA`, `BSW`, `BUC`, `LKA`, `LW`, `NV`, `SVC`, and `WAMC`. It targets `SME Generic System Name`. New filename aliases are accepted and converted back to the applicable old rows.
- **New** shows `BLS`, `BUC`, `FCR`, `FLS`, `FRS`, `LLS`, `LW`, `NV`, `PDS`, `RLS`, `RRS`, `SVC`, `WAMC`, `WSC`, and `WSR`, together with their component descriptions. It targets `Protech Generic System Name` directly and does not expand the selection into old acronyms.

Repair mode disables the Old/New toggle because Repair uses its own system list and matching rules.

| Older SME acronym/row | New Protech filename/row | Routing behavior |
| --- | --- | --- |
| `ACC`, `AEB` | `FRS` | ACC/AEB selections accept FRS. One FRS document can populate both rows when both intentionally map to FRS. |
| `FCW` | `FCR` | Either form is accepted for the family. |
| `APA` | `PDS` | Either form is accepted for the family. |
| `BSW`, `BSM` | `RRS` | These names are aliases for selection and placement. |
| `LKA` | `WSC` | Either form is accepted for the family. |
| `BUC` | `BUC` | Shared acronym; SME is checked first, then Protech. |
| `NV` | `NV` | Shared acronym; SME is checked first, then Protech. |
| `SVC` | `SVC` | Shared acronym; SME is checked first, then Protech. |

`AHL`, `LW`, and `WAMC` remain recognized without a separate old/new conversion. `BLS`, `FLS`, `LLS`, `RLS`, and `WSR` are new-only component acronyms.

The toggle decides which system column is authoritative:

- Old mode requires `SME Generic System Name` (or a recognized legacy system header).
- New mode requires `Protech Generic System Name`.

Hyper stops with a clear error if the workbook does not contain the header required by the selected mode. It does not silently switch formats.

## Repair systems

Repair mode uses Repair SharePoint roots and the Repair Systems checklist. It recognizes acronym modules including:

`SAS`, `YAW`, `G-Force`, `SWS`, `AHL`, `NV`, `HUD`, `SRS`, `SRA`, `ESC`, `SRS D&E`, `SCI`, `SRR`, `HLI`, `TPMS`, `SBI`, `RC`, `EBDE (1)`, `EBDE (2)`, `HDE (1)`, `HDE (2)`, `LGR`, `PSI`, `WRL`, `PCM`, `TRANS`, `AIR`, `ABS`, `BCM`, `ODS`, `OCS`, `OCS2`, `OCS3`, `OCS4`, `KEY`, `FOB`, `HVAC (1)`, `HVAC (2)`, `COOL`, `HEAD (1)`, and `HEAD (2)`.

Descriptive repair names are normalized before matching:

| Descriptive name | Acronym |
| --- | --- |
| Steering Angle Sensor | `SAS` |
| Yaw Rate Sensor | `YAW` |
| G Force Sensor | `G-Force` |
| Seat Weight Sensor | `SWS` |
| Adaptive Head Lamps | `AHL` |
| Night Vision | `NV` |
| Heads Up Display | `HUD` |
| Electronic Stability Control Relearn | `ESC` |
| Airbag Disengagement/Engagement | `SRS D&E` |
| Steering Column Inspection | `SCI` |
| Steering Rack Relearn | `SRR` |
| Headlamp Initialization | `HLI` |
| Tire Pressure Monitor Relearn | `TPMS` |
| Seat Belt Inspection | `SBI` |
| Battery Disengagement / Engagement | `EBDE (1)` / `EBDE (2)` |
| Hybrid Disengagement / Engagement | `HDE (1)` / `HDE (2)` |
| Liftgate Relearn | `LGR` |
| Power Seat Initialization | `PSI` |
| Window Relearn | `WRL` |
| Powertrain / Transmission module programming | `PCM` / `TRANS` |
| Airbag / Antilock Brake / Body module programming | `AIR` / `ABS` / `BCM` |
| Key Program / Key FOB Relearn | `KEY` / `FOB` |
| HVAC EVAC / Recharge | `HVAC (1)` / `HVAC (2)` |
| Coolant Services | `COOL` |
| Headset Reset, Spring / Squib Style | `HEAD (1)` / `HEAD (2)` |

Repair mode does not use the ADAS old/new SME-to-Protech routing. It matches the normalized repair module selected or extracted from the filename.

## Workbook requirements

Hyper prefers the worksheet `ADAS Model Version`; if it is absent, it uses the active sheet. Row 1 must contain recognizable headers. Comparison is case-insensitive and collapses extra spaces.

Required identifying columns:

- `Year`
- `Make`
- `Model`
- A system column: `SME Generic System Name`, `Generic System Name`, `System Name`, `System`, `Protech Generic System Name`, or `Protech Generic System`

Recognized hyperlink headers include `Service Information`, `Service Information Hyperlink`, `Service Information (URL)`, `Hyperlink`, `Link`, `URL`, `SI`, `SI Link`, and `SI URL`.

If no recognized hyperlink header exists, Hyper appends a new column. The normal population path creates `Service Information Hyperlink`.

For ADAS workbooks, Hyper also finds or appends a final `Calibration Type` column. When a new-format filename contains a recognized calibration value, that value is written to every exact vehicle/system row receiving the hyperlink. Old filenames without calibration metadata leave the cell blank.

Placeholder rows are excluded from the normal row index. Existing cells whose visible value is an HTTP URL but whose Excel hyperlink object points elsewhere are repaired so the target matches the visible URL.

## How Hyper places hyperlinks

Hyper builds an index keyed by normalized `(Year, Make, Model, System)` and uses this priority:

1. **Manual exception:** a filename in `SPECIFIC_HYPERLINKS` can target a fixed cell for a known workbook problem.
2. **ADAS header-aware match:** use the Old/New toggle to select the SME or Protech column, then match year, make, model, and system. If several rows intentionally share an acronym—or a new acronym maps to several old rows—write to all applicable rows for that exact vehicle.
3. **Exact indexed match:** exact year, make, raw model, and normalized system.
4. **Loose system match:** same year/make/model while ignoring a numeric system suffix if necessary.
5. **Regex model match:** strict year, make, and system with normalized model formatting/qualifiers.
6. **Fuzzy model match:** strict year, make, and system with a model similarity score of at least `0.72`.
7. **Missing-system rescue:** on an exact year/make/model row, inspect the applicable acronym column and use a safe model candidate.
8. **Letters-only system fallback:** allow variants such as `APA1` and `APA2` to share a base acronym and repeat model matching.
9. **Fallback row:** if nothing safe matches, append a row at the bottom rather than overwrite an unrelated vehicle.

An acronym verifier performs a final ADAS check. If a candidate row does not contain the filename's acronym, Hyper searches the same year/make/model rows and retargets the write to the correct system row.

Guardrails prevent fuzzy matches across unrelated model families. Certain ambiguous vehicles, including medium-duty Silverado/F-Series variants and explicit year/make/model exceptions, are forced to the bottom unless an exact match exists.

### Placement colors

| Appearance | Meaning |
| --- | --- |
| Blue, underlined | Exact placement. |
| Red, underlined | Approximate/relaxed or debug placement; review the row. |
| Dark yellow, underlined | A “No feature/document” special case. |
| Red text without a link | No usable SharePoint URL or no safe linked placement. |

Normally, the hyperlink cell displays the SharePoint URL and targets the same URL. Debug-writing can display friendly `Link For: ...` text instead.

For a fallback row, Hyper records the filename in the mode-specific error column. When possible, it also replaces an empty or `Placeholder` cell immediately left of the hyperlink with the PDF base name. These fallback values are red for manual review.

## Using Hyper

1. Start `Hyper.py` and complete the Hyper sign-in.
2. Select the Excel workbook(s).
3. Check manufacturer(s) in the same order as the workbooks.
4. Select one or more year ranges.
5. Choose **ADAS SI** or **Repair SI** with the mode switch.
6. For ADAS, choose **Old** or **New** and confirm that the system checklist changes to the expected acronym set.
7. Select the required systems.
8. Leave Broken Hyperlink Mode and Upload Mode clear for normal population.
9. Click **Start Automation** and verify that the confirmation summary shows the intended acronym mode.
10. Avoid interacting with the automated SharePoint page, clipboard, or automation profile while it runs.
11. Review terminal output, progress bars, edited workbooks, hyperlink rows, and the `Calibration Type` column.

Pause suspends the extractor and its child processes. Stop ends the current automation and resets the UI. Failed manufacturers are retried later in the batch, up to ten attempts.

## Broken Hyperlink Mode

Broken Hyperlink Mode uses two phases:

1. Index workbook rows containing real HTTP links, skipping blanks, `Placeholder`, `Hyperlink Not Available`, and non-URL text.
2. Open each link and check for SharePoint error panels and a loaded filename/viewer.
3. Record the year, make, model, and system of broken rows.
4. Search the applicable SharePoint year roots for replacement files.
5. Apply replacements without rescanning during the apply phase.

Cleanup combines selected year roots into one extractor process per manufacturer and narrows the search to years containing broken links when possible. The broken workbook entries, rather than normal system selections, drive the repair search. Always review results because Microsoft session state, SharePoint viewer behavior, and filenames affect classification.

## Upload Mode

Upload Mode is separate from hyperlink placement. It uses the ADAS/Repair mode, manufacturers, and year ranges to build jobs from machine-specific local roots in `Hyper.py`. Normal workbook/system inputs are not used for placement.

Before upload, `pdf_annotation_extractor.py`:

- scans local year folders recursively;
- extracts annotated pages;
- combines multipart PDFs;
- ignores Glass Statements, “No feature” documents, and configured support/job-aid documents;
- builds a processed tree under `%USERPROFILE%\Documents\Hyper Upload Cache`;
- creates an Excel report under that job's `_reports` folder; and
- uploads processed files while mirroring the folder tree.

The configured roots currently point to a specific workstation/user path and must be changed for another account or computer. The 2027–2031 destinations also need real URLs.

## Logs and reports

Runtime logs are written to:

```text
%USERPROFILE%\Documents\Hyper Logs\Hyper_Log_MM_DD_YYYY_HH-MM-SS.log
```

After a normal batch, Hyper writes a summary in the same folder:

```text
ADAS_SI_Report_YYYY-MM-DD_HH-MM-SS.txt
```

Despite the filename prefix, its header changes between ADAS SI and Repair SI. It contains combined runtime/file totals and per-manufacturer totals for each year-range link.

Useful messages:

- `ADAS header match` — an SME/Protech row was resolved directly.
- `Acronym verifier: retarget` — the write moved to the correct system row.
- `[exact]` — exact placement.
- `[approx]` — relaxed placement; review the red link.
- `No hyperlink ... adding ... as placeholder` — a red review entry was created.
- `Finished SharePoint link ...` — a year-range root completed.

## Troubleshooting

- **Hyper crashes during startup:** first confirm that Google Chrome is fully updated. Hyper requires a compatible Chrome and ChromeDriver version, and an outdated or partially updated Chrome installation is the most likely cause of an immediate startup failure. Complete any pending Chrome update, close and reopen Chrome, and then restart Hyper.
- **SharePoint never loads:** verify the real Chrome profile is signed in and the copied automation profile has a current session and target-library access.
- **ChromeDriver mismatch:** Hyper checks Chrome, removes a mismatched cached driver, and uses `chromedriver-autoinstaller`; network access may be required if the driver is not cached.
- **Wrong workbook/make pairing:** reselect files and manufacturers in matching order.
- **Rows are skipped:** check row-one headers and ensure the intended row is not `Placeholder` in the hyperlink column.
- **Files are skipped:** ensure the filename has a four-digit year, correct make/model, and recognized acronym.
- **Many red links:** naming differences caused relaxed or fallback placement; review before distributing the workbook.
- **Wrong ADAS column:** confirm whether the filename uses the older SME or newer Protech acronym and that the corresponding header exists.
- **SharePoint list appears incomplete:** do not interact with the automated browser; changing the page can stale Selenium rows and force a restart.
- **Upload finds no PDFs:** verify the machine-specific local root, selected years, filters, and preprocessing output.

## Current limitations

- Hyper is Windows-specific and depends on Chrome, registry/profile paths, the clipboard, and Windows automation libraries.
- SharePoint URLs and upload roots are embedded in source.
- Hyper and SharePoint authentication are not centrally managed.
- Workbooks are modified in place rather than written to a separate output.
- Filename and header conventions are part of the data contract.
- Some workbook inconsistencies use explicit special cases and manual mappings in `SharepointExtractor.py`.
- 2027–2031 configuration is incomplete.
- No committed `requirements.txt` currently reproduces the environment.

## Maintenance notes

When adding a new system or naming convention, update all relevant layers together:

1. the UI checklist in `Hyper.py`;
2. ADAS alias expansion or Repair synonym normalization;
3. the extractor's known/defined module names;
4. system-column routing for old/new ADAS acronyms;
5. the SharePoint filename convention; and
6. workbook rows or a narrowly scoped manual exception.

Test with a copy of a representative workbook. Confirm exact, approximate, missing-link, multi-row alias, and no-document cases before production use.
