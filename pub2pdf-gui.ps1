#Requires -Version 5.1
<#
.SYNOPSIS
    Simple window for converting Publisher .pub files to PDF.

.DESCRIPTION
    Teacher-friendly launcher. Pick a folder, click Convert - no editing required.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConverterScript = Join-Path $ScriptDir "pub2pdf.ps1"

if (-not (Test-Path -LiteralPath $ConverterScript)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Could not find pub2pdf.ps1 in:`n$ScriptDir`n`nPlease keep all files from the download together in one folder.",
        "pub2pdf",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Publisher to PDF"
$form.Size = New-Object System.Drawing.Size(640, 520)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(560, 460)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$iconPath = Join-Path $ScriptDir "pub2pdf.ico"
if (Test-Path -LiteralPath $iconPath) {
    $form.Icon = New-Object System.Drawing.Icon($iconPath)
}

$intro = New-Object System.Windows.Forms.Label
$intro.Location = New-Object System.Drawing.Point(16, 12)
$intro.Size = New-Object System.Drawing.Size(590, 48)
$intro.Text = "Choose the folder that contains your Publisher files. Every .pub file in that folder and its subfolders will be converted to PDF."
$form.Controls.Add($intro)

$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Location = New-Object System.Drawing.Point(16, 68)
$sourceLabel.Size = New-Object System.Drawing.Size(300, 24)
$sourceLabel.Text = "Folder with your .pub files:"
$form.Controls.Add($sourceLabel)

$sourceBox = New-Object System.Windows.Forms.TextBox
$sourceBox.Location = New-Object System.Drawing.Point(16, 92)
$sourceBox.Size = New-Object System.Drawing.Size(500, 24)
$form.Controls.Add($sourceBox)

$sourceBrowse = New-Object System.Windows.Forms.Button
$sourceBrowse.Location = New-Object System.Drawing.Point(524, 90)
$sourceBrowse.Size = New-Object System.Drawing.Size(84, 28)
$sourceBrowse.Text = "Browse..."
$form.Controls.Add($sourceBrowse)

$sameFolderCheck = New-Object System.Windows.Forms.CheckBox
$sameFolderCheck.Location = New-Object System.Drawing.Point(16, 128)
$sameFolderCheck.Size = New-Object System.Drawing.Size(580, 24)
$sameFolderCheck.Text = "Save each PDF in the same folder as its .pub file (recommended)"
$sameFolderCheck.Checked = $true
$form.Controls.Add($sameFolderCheck)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(16, 160)
$outputLabel.Size = New-Object System.Drawing.Size(300, 24)
$outputLabel.Text = "Or save all PDFs into this folder instead:"
$outputLabel.Enabled = $false
$form.Controls.Add($outputLabel)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(16, 184)
$outputBox.Size = New-Object System.Drawing.Size(500, 24)
$outputBox.Enabled = $false
$form.Controls.Add($outputBox)

$outputBrowse = New-Object System.Windows.Forms.Button
$outputBrowse.Location = New-Object System.Drawing.Point(524, 182)
$outputBrowse.Size = New-Object System.Drawing.Size(84, 28)
$outputBrowse.Text = "Browse..."
$outputBrowse.Enabled = $false
$form.Controls.Add($outputBrowse)

$skipCheck = New-Object System.Windows.Forms.CheckBox
$skipCheck.Location = New-Object System.Drawing.Point(16, 220)
$skipCheck.Size = New-Object System.Drawing.Size(580, 24)
$skipCheck.Text = "Skip files that already have a PDF (safe to run again)"
$skipCheck.Checked = $true
$form.Controls.Add($skipCheck)

$overwriteCheck = New-Object System.Windows.Forms.CheckBox
$overwriteCheck.Location = New-Object System.Drawing.Point(16, 246)
$overwriteCheck.Size = New-Object System.Drawing.Size(580, 24)
$overwriteCheck.Text = "Replace existing PDFs"
$form.Controls.Add($overwriteCheck)

$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Location = New-Object System.Drawing.Point(16, 282)
$convertButton.Size = New-Object System.Drawing.Size(160, 36)
$convertButton.Text = "Convert"
$form.Controls.Add($convertButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(188, 290)
$statusLabel.Size = New-Object System.Drawing.Size(420, 24)
$statusLabel.Text = "Close Microsoft Publisher before you start."
$form.Controls.Add($statusLabel)

$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Location = New-Object System.Drawing.Point(16, 326)
$logLabel.Size = New-Object System.Drawing.Size(200, 24)
$logLabel.Text = "Progress:"
$form.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(16, 350)
$logBox.Size = New-Object System.Drawing.Size(592, 110)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

function Set-UiBusy {
    param([bool]$Busy)

    $convertButton.Enabled = -not $Busy
    $sourceBrowse.Enabled = -not $Busy
    $outputBrowse.Enabled = (-not $Busy) -and (-not $sameFolderCheck.Checked)
    $sourceBox.Enabled = -not $Busy
    $outputBox.Enabled = (-not $Busy) -and (-not $sameFolderCheck.Checked)
    $sameFolderCheck.Enabled = -not $Busy
    $skipCheck.Enabled = -not $Busy
    $overwriteCheck.Enabled = -not $Busy
}

function Add-LogLine {
    param([string]$Line)

    if ($logBox.Text.Length -gt 0) {
        $logBox.AppendText("`r`n")
    }

    $logBox.AppendText($Line)
    [System.Windows.Forms.Application]::DoEvents()
}

$sameFolderCheck.Add_CheckedChanged({
    $enabled = -not $sameFolderCheck.Checked
    $outputLabel.Enabled = $enabled
    $outputBox.Enabled = $enabled
    $outputBrowse.Enabled = $enabled
})

$skipCheck.Add_CheckedChanged({
    if ($skipCheck.Checked) {
        $overwriteCheck.Checked = $false
    }
})

$overwriteCheck.Add_CheckedChanged({
    if ($overwriteCheck.Checked) {
        $skipCheck.Checked = $false
    }
})

$sourceBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select the folder that contains your .pub files"
    $dialog.ShowNewFolderButton = $false

    if ($sourceBox.Text -and (Test-Path -LiteralPath $sourceBox.Text)) {
        $dialog.SelectedPath = $sourceBox.Text
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sourceBox.Text = $dialog.SelectedPath
    }
})

$outputBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select where PDF files should be saved"
    $dialog.ShowNewFolderButton = $true

    if ($outputBox.Text -and (Test-Path -LiteralPath $outputBox.Text)) {
        $dialog.SelectedPath = $outputBox.Text
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputBox.Text = $dialog.SelectedPath
    }
})

$convertButton.Add_Click({
    $source = $sourceBox.Text.Trim()
    if (-not $source) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please choose the folder that contains your .pub files.",
            "pub2pdf",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $source)) {
        [System.Windows.Forms.MessageBox]::Show(
            "That folder was not found:`n$source",
            "pub2pdf",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ($sameFolderCheck.Checked) {
        $output = $source
    }
    else {
        $output = $outputBox.Text.Trim()
        if (-not $output) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please choose where PDF files should be saved, or tick the box to save them beside each .pub file.",
                "pub2pdf",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }
    }

    $pubFiles = @(Get-ChildItem -LiteralPath $source -Filter "*.pub" -Recurse -File -ErrorAction Stop)
    if ($pubFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No .pub files were found in that folder or its subfolders.",
            "pub2pdf",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Found $($pubFiles.Count) .pub file(s).`n`nPublisher will open automatically. Each file usually takes 30-120 seconds.`n`nContinue?",
        "pub2pdf",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $logBox.Clear()
    Set-UiBusy -Busy $true
    $statusLabel.Text = "Converting - please wait..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        $params = @{
            SourceRoot = $source
            OutputRoot = $output
            LogCallback = { param($Line) Add-LogLine -Line $Line }
        }

        if ($skipCheck.Checked) {
            $params.Skip = $true
        }
        elseif ($overwriteCheck.Checked) {
            $params.Overwrite = $true
        }

        & $ConverterScript @params

        $statusLabel.Text = "Finished. See the log below for details."
        [System.Windows.Forms.MessageBox]::Show(
            "Conversion finished.`n`nCheck the progress box for any files that failed.",
            "pub2pdf",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Add-LogLine -Line "ERROR: $($_.Exception.Message)"
        $statusLabel.Text = "Something went wrong."
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "pub2pdf",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Set-UiBusy -Busy $false
    }
})

[void]$form.ShowDialog()
