# pub2pdf — Installation and use

Bulk-convert **Microsoft Publisher** files (`.pub`) to **PDF**. No coding required.

Project page: [github.com/spufidoo/pub2pdf](https://github.com/spufidoo/pub2pdf)

---

## What you need first

- A **Windows** laptop or PC
- **Microsoft Publisher** installed (the same program you use to open `.pub` files)
- **PowerShell** — already built into Windows; the `.bat` file uses `powershell.exe`, which is correct
- Your `.pub` files on the computer — if they are in **OneDrive**, see [OneDrive files](#onedrive-files) below
- Permission to run scripts — if your workplace blocks them, see [Step 4 (alternative)](#step-4-alternative-if-your-computer-blocks-the-tool)

---

## Step 1: Download the tool

1. Open [github.com/spufidoo/pub2pdf](https://github.com/spufidoo/pub2pdf) in your web browser.
2. Click the green **Code** button (top right).
3. Click **Download ZIP**.
4. Open your **Downloads** folder and double-click the ZIP file.
5. Copy these files into a folder you can find easily (for example **Documents → pub2pdf**):
   - `pub2pdf.ps1`
   - `Convert my Publisher files.bat`

You do **not** need a GitHub account.

If the ZIP creates a folder called `pub2pdf-main` or `pub2pdf.ps1-main`, open that folder — the files are inside it.

---

## Step 2: Unblock the files (important)

Windows may block files downloaded from the internet.

1. Open the folder where you saved the files
2. **Right-click** `pub2pdf.ps1` → **Properties**
3. If you see **Unblock** at the bottom, tick it → **OK**
4. Repeat for `Convert my Publisher files.bat` if it also has an **Unblock** box

---

## Step 3: Edit the “Run” file (double-click to start)

Open **`Convert my Publisher files.bat`** in Notepad (right-click → **Edit**).

Change only the **`SOURCE`** and **`OUTPUT`** lines at the top. Save and close Notepad.

**Important:** Do not split the `powershell.exe` command across two lines.

### Example — separate folders

If your Publisher files are here:

`C:\Users\Jane\OneDrive - School\Year 3\Plans`

and you want PDFs here:

`C:\Users\Jane\OneDrive - School\Year 3\PDFs`

edit the `.bat` file to:

```bat
REM === EDIT THESE TWO FOLDERS ===
set "SOURCE=C:\Users\Jane\OneDrive - School\Year 3\Plans"
set "OUTPUT=C:\Users\Jane\OneDrive - School\Year 3\PDFs"
```

### Advanced example — PDF beside each `.pub`

If you have many `.pub` files in a folder tree, put **`pub2pdf.ps1`** and the **`.bat`** file in the **top folder** that contains them, then set:

```bat
REM === EDIT THESE TWO FOLDERS ===
set "SOURCE=."
set "OUTPUT=."
```

**`.`** means “this folder” (where the `.bat` file lives). The tool searches subfolders for `.pub` files and creates matching PDFs alongside them.

**Tip:** To copy a folder path without typing it — open the folder in File Explorer, click the address bar, press Ctrl+C, paste into Notepad.

**Remember:**

- **`SOURCE`** = folder containing your `.pub` files (searches subfolders too)
- **`OUTPUT`** = where PDFs are saved, keeping the same subfolder structure
- **`SOURCE`** and **`OUTPUT`** can be the same (use `.` for both) if you want PDFs beside each `.pub`

### Optional — longer export timeout

If export seems slow or stalls, uncomment and edit this line in the `.bat` file:

```bat
REM set "TIMEOUT=600"
```

Change to (600 seconds = 10 minutes):

```bat
set "TIMEOUT=600"
```

Leave it commented out to use the default (180 seconds). **Do not** glue `600` to other text on the same line.

### Or build your own `.bat` from scratch

<details>
<summary>Click to expand a full template</summary>

```bat
@echo off
cd /d "%~dp0"

REM === EDIT THESE TWO FOLDERS ===
set "SOURCE=PUT YOUR PUBLISHER FOLDER HERE"
set "OUTPUT=PUT YOUR PDF FOLDER HERE"

REM Optional: seconds to wait for export (600 = 10 minutes).
REM Leave commented out to use the script default (180).
REM set "TIMEOUT=600"

if defined TIMEOUT (
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "%SOURCE%" -OutputRoot "%OUTPUT%" -Skip -ExportTimeoutSeconds %TIMEOUT%
) else (
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "%SOURCE%" -OutputRoot "%OUTPUT%" -Skip
)

echo.
echo Finished. Press any key to close this window.
pause >nul
```

</details>

---

## Step 4: Run the tool

1. Close **Microsoft Publisher** if it is open
2. Double-click **`Convert my Publisher files.bat`**
3. A dark or blue window shows progress — leave it open until **Conversion complete**
4. PDFs appear in your **OUTPUT** folder (mirroring subfolders from **SOURCE**)

### Log file

The tool writes a log file named `PublisherConversion_....log`.

- It tries your **OUTPUT** folder first
- On OneDrive, it may fall back to your **Temp** folder instead

Look for this line near the start of the run:

```text
Log    : C:\Users\...\PublisherConversion_....log
```

---

## Step 4 (alternative): If your computer blocks the tool

You may see “running scripts is disabled” or “execution policy”.

**Try this first:** the `.bat` file already uses `-ExecutionPolicy Bypass`.

**If that is not enough:**

1. Click **Start**, type **PowerShell**, click **Open**
2. At the `PS C:\Users\...>` prompt, type:

   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. Press Enter, then run the `.bat` file again

**If it still fails:** contact **IT support** and show them this page. Do not change other security settings unless IT tells you to.

---

## What the options mean

The `.bat` file passes these switches to the script:

| Switch | What it does |
|--------|----------------|
| `-Skip` | **Recommended.** Skips files that already have a PDF — use on repeat runs |
| `-Overwrite` | Converts every file again and replaces existing PDFs |

**Do not use `-Skip` and `-Overwrite` together.**

---

## OneDrive files

OneDrive can block new files until folders are synced locally.

1. In File Explorer, **right-click** the folder with your `.pub` files
2. Choose **Always keep on this device**
3. Wait for green tick icons on the files
4. Run the tool

This avoids “cannot open file”, log, and PDF save errors.

If PDFs still fail to save into OneDrive, set **`OUTPUT`** to a local folder (for example `C:\Temp\PDFs`) and move the PDFs afterwards.

---

## Test with one file first (optional)

Ask IT or a colleague to change the `.bat` to use `-File` instead of `-SourceRoot` for a single test file. Most people use the folder method above.

---

## If something goes wrong

| Problem | What to try |
|---------|-------------|
| “Publisher cannot open the file” | Open that `.pub` manually in Publisher — if Publisher fails too, the file may be damaged or not a real Publisher file |
| Nothing happens when you double-click the `.bat` | Ensure `pub2pdf.ps1` is in the **same folder** as the `.bat` file |
| “Running scripts is disabled” | See [Step 4 (alternative)](#step-4-alternative-if-your-computer-blocks-the-tool) |
| Tool stops after `TEMP` with no progress | Check the **taskbar for Publisher** — a dialog may be waiting; raise `TIMEOUT` if needed |
| “600powershell.exe” or “Cannot convert … ExportTimeoutSeconds” | Typo in `.bat` — use the template above; `-ExportTimeoutSeconds 600` must be separate from other text |
| “GetRelativePath” error | Download the latest `pub2pdf.ps1` from GitHub |
| “Could not find file … PublisherConversion…log” | OneDrive blocked the log — use latest script; log path is shown as `Log : ...` (often in Temp) |
| Export OK but “Could not find file … .pdf” on copy | OneDrive blocked the PDF — **Always keep on this device**, latest script, or use a local `OUTPUT` folder |
| Tool stops after many files | Run again with `-Skip` — already-converted files are skipped |

---

## Folder layout (minimum)

```
Documents
└── pub2pdf
    ├── pub2pdf.ps1
    └── Convert my Publisher files.bat
```

Your `.pub` files can live anywhere — set **`SOURCE`** and **`OUTPUT`** in the `.bat` file.

---

## Need help?

### Ask IT or a colleague

Send them:

1. The **log file** (path shown as `Log : ...` in the run window)
2. A screenshot of any error message
3. The name of one `.pub` file that failed

### Raise a GitHub issue (optional)

If you want to report a bug or ask for help on the project itself:

1. Open [github.com/spufidoo/pub2pdf](https://github.com/spufidoo/pub2pdf) in your web browser
2. Click the **Issues** tab (near the top)
3. Click the green **New issue** button
4. Give it a short title (for example: “PDF not created in OneDrive folder”)
5. In the description, include:
   - What you were trying to do
   - The **log file** contents or the error message (copy and paste from the black/blue window)
   - Your Windows version and Publisher version if you know them
6. Click **Submit new issue**

You need a **free GitHub account** to raise an issue. If you do not have one, click **Sign up** on GitHub first, or ask a colleague who has an account to raise it for you.
