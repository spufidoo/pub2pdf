# pub2pdf — Installation and use

This tool turns **Microsoft Publisher** files (`.pub`) into **PDF** files. You do not need to know how to code. Follow the steps below in order.

---

## What you need first

Check all of these before you start:

- A **Windows** laptop or PC
- **Microsoft Publisher** installed (the same program you use to open `.pub` files)
- **PowerShell** — already built into Windows; you do not need to install anything extra. The `.bat` file uses the normal Windows version (`powershell.exe`), which is correct
- Your `.pub` files saved on the computer — if they are in **OneDrive**, make sure they are downloaded locally (see [OneDrive tip](#onedrive-files) below)
- Permission to run scripts — if your school or workplace blocks them, ask **IT support** to help with [Step 4](#step-4-if-your-computer-blocks-the-tool)

---

## Step 1: Download the tool
<!--
You only need **one file**: `pub2pdf.ps1`. You will find the file here (https://github.com/spufidoo/pub2pdf/blob/main/pub2pdf.ps1). Click on the download icon on the header bar to download it.
<img width="905" height="51" alt="image" src="https://github.com/user-attachments/assets/f664e7d3-2028-4022-b17a-1f6367c1a907" />
-->

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

## Step 3: Edit the “Run” file (double-click to start)

You will create a small helper file so you never have to type commands yourself.

### Easy option: use the template file

1. Right-click the `Convert my Publisher files.bat` file → **Edit** (or open it in Notepad)
2. Change only the two folder paths at the top (`SOURCE` and `OUTPUT`)
3. Save and close Notepad

**Important:** The `powershell.exe` line must be **one single line** in Notepad — do not split it across two lines unless you know batch files well.

### Example paths

If your Publisher files are here:

`C:\Users\Jane\OneDrive - School\Year 3\Plans`

and you want PDFs here:

`C:\Users\Jane\OneDrive - School\Year 3\PDFs`

then you should edit these lines:

```bat
REM === EDIT THESE TWO FOLDERS ===
set "SOURCE=C:\Users\Jane\OneDrive - School\Year 3\Plans"
set "OUTPUT=C:\Users\Jane\OneDrive - School\Year 3\PDFs"
```
### Another (Advanced) Example

If you want your PDFs to be in the same folder as your PUBs, and you happen to have MANY Publisher documents, then copy the two programs to the top folder (e.g. Documents), edit the `Convert my Publisher files.bat` file by changing the following lines:
```bat
REM === EDIT THESE TWO FOLDERS ===
set "SOURCE=."
set "OUTPUT=."
```
This means "wherever you find a Publisher file, create a PDF of it along side it."

**Tip:** To copy a folder path without typing it:

1. Open File Explorer and go to the folder
2. Click the address bar at the top
3. Copy the path (Ctrl+C)
4. Paste it between the quote marks in Notepad

**Important:**

- **`-SourceRoot`** = the folder that contains your `.pub` files
- **`-OutputRoot`** = a **different** folder where you want the PDFs saved (for example, create a folder called `PDFs`). Do not use the same folder as the tool itself unless you deliberately want PDFs saved beside the script

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

1. Open a PowerShell window:
   Click the Start button (bottom left of your Start bar), and start typing ``PowerShell``. Click Open.
2. After the ``PS C:\Users\YourName>`` prompt, type
   ``Set-ExecutionPolicy RemoteSigned -Scope CurrentUser``
   and hit Enter.

If that doesn't work, ask your IT Support.

---

## What the options mean

You can change the last part of the line in your `.bat` file:

| Add this | What it does |
|----------|----------------|
| `-Skip` | **Recommended.** Skips files that already have a PDF — use this when you run the tool again |
| `-Overwrite` | Converts every file again and replaces existing PDFs |
| *(neither)* | Converts every file; may overwrite PDFs if they already exist |

**Do not use `-Skip` and `-Overwrite` together.**

### Longer export timeout (optional)

If export seems to hang, you can allow more time (in seconds). Add this to the **end** of the same `powershell.exe` line, with a **space** before it:

```bat
-Skip -ExportTimeoutSeconds 600
```

The full line must look like one command — for example:

```bat
powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "C:\Your\Pub folder" -OutputRoot "C:\Your\PDF folder" -Skip -ExportTimeoutSeconds 600
```

There must be a **space** after `600`. Do not type `600` right next to `powershell.exe` or any other text.

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
| Tool stops after `TEMP` with no progress | Check the **taskbar for Publisher** — a dialog may be waiting. After 3 minutes the tool stops automatically unless you raised the timeout |
| “600powershell.exe” or “Cannot convert … ExportTimeoutSeconds” | Your `.bat` file has a typo — the timeout value is stuck to other text. Use one `powershell.exe` line with a space: `-ExportTimeoutSeconds 600` |
| “GetRelativePath” error | Copy the latest `pub2pdf.ps1` from the project — you may have an older version |
| PDFs saved in the wrong place | Check **`-OutputRoot`** in your `.bat` file points at a PDF folder, not the script folder |
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
