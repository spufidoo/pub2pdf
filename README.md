# pub2pdf

Bulk-convert **Microsoft Publisher** files (`.pub`) to **PDF** on Windows.

Project page: [github.com/spufidoo/pub2pdf](https://github.com/spufidoo/pub2pdf)

## Background

Microsoft is retiring Publisher in October 2026. This tool helps preserve existing `.pub` documents by converting them to PDF before they become difficult to open.

## Requirements

- Windows PC with **Microsoft Publisher** installed
- **PowerShell** (included with Windows)
- `.pub` files available locally (see [OneDrive](#onedrive) if files are cloud-synced)

## Quick start (recommended)

1. Download **`pub2pdf.zip`** from [releases/latest](https://github.com/spufidoo/pub2pdf/releases/latest).
2. **Extract** the ZIP (right-click → **Extract All…** → choose a folder such as `Documents\pub2pdf`).
3. Double-click **`Convert Publisher to PDF.bat`**.

A small window titled **Publisher to PDF** opens. Click **Browse…**, choose the folder that contains your `.pub` files, then click **Convert**. Every `.pub` file in that folder and its subfolders is converted. PDFs are saved beside each source file by default.

Close Microsoft Publisher before you start. The window shows progress and writes a log file when it finishes.

No editing, unblocking, or technical setup is required — the `.bat` file handles that for you.

### Using the window

| Option | What it does |
|--------|----------------|
| **Save each PDF in the same folder as its .pub file** | Recommended. Leaves each PDF next to the file it came from. |
| **Or save all PDFs into this folder instead** | Uncheck the option above, then browse to one output folder (keeps subfolder structure). |
| **Skip files that already have a PDF** | Default. Safe to run again — only missing PDFs are created. |
| **Replace existing PDFs** | Re-converts every file and overwrites PDFs that are already there. Do not use with **Skip** ticked. |

After you click **Convert**, you are asked to confirm how many `.pub` files were found. Each file usually takes 30–120 seconds while Publisher works in the background.

## Advanced: batch file for fixed folders

For repeat runs with fixed folders, edit **`Convert my Publisher files.bat`** in Notepad:

```bat
set "SOURCE=C:\path\to\your\pub\files"
set "OUTPUT=C:\path\to\save\pdfs"
```

Use `.` for both to place each PDF beside its `.pub` file. Close Publisher, then double-click the `.bat` file.

## Options (advanced)

These apply if you run `pub2pdf.ps1` from PowerShell or edit **`Convert my Publisher files.bat`**:

| Switch | Purpose |
|--------|---------|
| `-Skip` | Skip files that already have a PDF (same as the GUI default) |
| `-Overwrite` | Re-convert and replace existing PDFs |

Do not use `-Skip` and `-Overwrite` together.

To increase the export timeout, uncomment `set "TIMEOUT=600"` in `Convert my Publisher files.bat` (default: 180 seconds).

## OneDrive

Right-click the source folder → **Always keep on this device**, wait for sync, then run. If saves fail, save PDFs to a local folder instead (uncheck “same folder” in the GUI, or set **OUTPUT** to e.g. `C:\Temp\PDFs` in the `.bat` file) and move PDFs afterwards.

## Troubleshooting

| Problem | Suggestion |
|---------|------------|
| Scripts disabled | The `.bat` files use `-ExecutionPolicy Bypass`. If blocked, run `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` in PowerShell, or contact IT. |
| Publisher cannot open file | Try opening the `.pub` manually — the file may be damaged. |
| Nothing happens | Ensure all files from the ZIP are in the same folder. |
| Stalls or times out | Check the taskbar for a Publisher dialog; increase `TIMEOUT` in the advanced `.bat` file. |
| OneDrive save errors | Use **Always keep on this device** or save PDFs to a local folder. |
| Stops after many files | Run again — “Skip existing PDFs” resumes where it left off. |

## Help

For bugs or feature requests, open an [issue](https://github.com/spufidoo/pub2pdf/issues) on GitHub. Include the log file path, any error message, and one example `.pub` filename.
