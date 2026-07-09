# pub2pdf — Installation and use

This tool turns **Microsoft Publisher** files (`.pub`) into **PDF** files. You do not need to know how to code. Follow the steps below in order.

---

## What you need first

Check all of these before you start:

- A **Windows** laptop or PC
- **Microsoft Publisher** installed (the same program you use to open `.pub` files)
- Your `.pub` files saved on the computer — if they are in **OneDrive**, make sure they are downloaded locally (see [OneDrive tip](#onedrive-files) below)
- Permission to run scripts — if your school or workplace blocks them, ask **IT support** to help with [Step 4](#step-4-if-your-computer-blocks-the-tool)

---

## Step 1: Download the tool

You only need **one file**: `pub2pdf.ps1`.

### If someone sent you the file

1. Save `pub2pdf.ps1` to a folder you can find easily, for example:  
   **Documents → pub2pdf**

### If you are downloading from GitHub

1. Open the project page in your web browser (your colleague will give you the link).
2. Click the green **Code** button near the top right.
3. Click **Download ZIP**.
4. Open your **Downloads** folder and double-click the ZIP file.
5. Drag **`pub2pdf.ps1`** into **Documents → pub2pdf** (create the `pub2pdf` folder if it does not exist).

You do **not** need to install GitHub or create an account.

---

## Step 2: Unblock the file (important)

Windows may block files downloaded from the internet.

1. Open **Documents → pub2pdf**
2. **Right-click** `pub2pdf.ps1`
3. Click **Properties**
4. At the bottom, if you see a tick box called **Unblock**, tick it
5. Click **OK**

---

## Step 3: Create a “Run” file (double-click to start)

You will create a small helper file so you never have to type commands yourself.

1. Open **Notepad** (search for it in the Start menu)
2. Copy **all** of the text inside the box below
3. Paste it into Notepad
4. Change the two folder paths on the lines that start with `-SourceRoot` and `-OutputRoot` to match **your** folders (see the example underneath)
5. Click **File → Save As**
6. Save in the **same folder** as `pub2pdf.ps1` (Documents → pub2pdf)
7. Set **Save as type** to **All files**
8. Name the file: **`Convert my Publisher files.bat`**
9. Click **Save**

### Text to copy into Notepad

```bat
@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "PUT YOUR PUBLISHER FOLDER HERE" -OutputRoot "PUT YOUR PDF FOLDER HERE" -Skip
echo.
echo Finished. Press any key to close this window.
pause >nul
```

### Example paths

If your Publisher files are here:

`C:\Users\Jane\OneDrive - School\Year 3\Plans`

and you want PDFs here:

`C:\Users\Jane\OneDrive - School\Year 3\PDFs`

then your file should look like:

```bat
@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "C:\Users\Jane\OneDrive - School\Year 3\Plans" -OutputRoot "C:\Users\Jane\OneDrive - School\Year 3\PDFs" -Skip
echo.
echo Finished. Press any key to close this window.
pause >nul
```

**Tip:** To copy a folder path without typing it:

1. Open File Explorer and go to the folder
2. Click the address bar at the top
3. Copy the path (Ctrl+C)
4. Paste it between the quote marks in Notepad

---

## Step 4: Run the tool

1. Make sure **Microsoft Publisher is closed**
2. Double-click **`Convert my Publisher files.bat`**
3. A dark or blue window will open and show progress — **leave it open** until it says **Conversion complete**
4. Your PDF files will be in the **OutputRoot** folder you chose (including subfolders)

The tool also writes a **log file** in that same output folder. The file name starts with `PublisherConversion_`. Open it in Notepad if you need to see what happened.

---

## Step 4 (alternative): If your computer blocks the tool

Some workplaces stop scripts from running. You may see a red error about “execution policy” or “running scripts is disabled”.

**Try this first:** Step 3 already includes a fix (`-ExecutionPolicy Bypass`). If it still fails:

1. Contact **IT support**
2. Show them this file and ask them to either:
   - Allow `pub2pdf.ps1` to run for your user account, or
   - Run the tool once for you using the `.bat` file from Step 3

Do not change security settings yourself unless IT tells you to.

---

## What the options mean

You can change the last part of the line in your `.bat` file:

| Add this | What it does |
|----------|----------------|
| `-Skip` | **Recommended.** Skips files that already have a PDF — use this when you run the tool again |
| `-Overwrite` | Converts every file again and replaces existing PDFs |
| *(neither)* | Converts every file; may overwrite PDFs if they already exist |

**Do not use `-Skip` and `-Overwrite` together.**

---

## OneDrive files

If your `.pub` files are in OneDrive:

1. In File Explorer, **right-click** the folder that contains your `.pub` files
2. Choose **Always keep on this device**
3. Wait until OneDrive finishes syncing (green tick on the files)
4. Then run the tool

This avoids “cannot open file” errors.

---

## Test with one file first (optional)

To try a single file before converting a whole folder, ask IT or a colleague to adjust your `.bat` file to use `-File` instead of `-SourceRoot`. This is optional — most people use the folder method in Step 3.

---

## If something goes wrong

| Problem | What to try |
|---------|-------------|
| “Publisher cannot open the file” | Open that `.pub` file manually in Publisher. If Publisher cannot open it either, the file may be damaged or not a real Publisher file |
| Nothing happens when you double-click the `.bat` file | Make sure `pub2pdf.ps1` is in the **same folder** as the `.bat` file |
| “Running scripts is disabled” | Ask IT support (see Step 4 alternative) |
| PDFs missing for some files | Open the log file in the output folder and look for lines starting with `FAILED` |
| Tool stops after many files | Run it again with `-Skip` — it will continue where it left off |

---

## Folder layout (what you should have)

```
Documents
└── pub2pdf
    ├── pub2pdf.ps1
    └── Convert my Publisher files.bat
```

Your Publisher files and PDF output can live anywhere — you choose those paths in the `.bat` file.

---

## Need help?

Send your IT support or colleague:

1. The **log file** from the output folder (`PublisherConversion_....log`)
2. A screenshot of any error message
3. The name of one `.pub` file that failed
