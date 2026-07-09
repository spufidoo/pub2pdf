    <#
    .SYNOPSIS
        Bulk or single-file convert Microsoft Publisher .pub files to PDF.

    .DESCRIPTION
        Uses Microsoft Publisher COM automation to open .pub files and export them
        as PDF using Document.ExportAsFixedFormat.

    .REQUIREMENTS
        - Windows
        - Microsoft Publisher installed and COM-registered
        - Windows PowerShell 5.1 or PowerShell 7

    .EXAMPLES
        Bulk convert a folder tree:

            .\pub2pdf.ps1 `
                -SourceRoot "C:\Users\User\Documents\Publisher" `
                -OutputRoot "C:\Temp\PubPDF"

        Test a single file:

            .\pub2pdf.ps1 `
                -File "C:\Temp\PubTest\Test.pub" `
                -OutputRoot "C:\Temp\PubTest" `
                -Overwrite

        Convert current folder tree into PDFs beside the source files:

            .\pub2pdf.ps1 -SourceRoot . -OutputRoot . -Overwrite

        Resume a bulk run, skipping files that already have a PDF:

            .\pub2pdf.ps1 -SourceRoot . -OutputRoot . -Skip
    #>

    [CmdletBinding(DefaultParameterSetName = "Bulk")]
    param(
        [Parameter(Mandatory, ParameterSetName = "Bulk")]
        [string]$SourceRoot,

        [Parameter(Mandatory, ParameterSetName = "Single")]
        [string]$File,

        [Parameter(Mandatory)]
        [string]$OutputRoot,

        [switch]$Overwrite,

        # Skip .pub files when the target PDF already exists.
        [switch]$Skip,

        # Restart Publisher every N opened files to avoid COM instability on long runs. 0 = never.
        [int]$RestartEvery = 50,

        # Stop waiting if PDF export takes longer than this (seconds). 0 = no timeout.
        [int]$ExportTimeoutSeconds = 180
    )

    if ($Skip -and $Overwrite) {
        throw "Use either -Skip or -Overwrite, not both."
    }

    $ErrorActionPreference = "Stop"

    # Publisher constants (PbFixedFormatType / PbFixedFormatIntent)
    $pbFixedFormatTypePDF = 2
    $pbIntentStandard = 2

    function Get-RelativePathCompat {
        param(
            [Parameter(Mandatory)]
            [string]$BasePath,

            [Parameter(Mandatory)]
            [string]$TargetPath
        )

        $BasePath = (Resolve-Path -LiteralPath $BasePath).Path.TrimEnd('\')
        $TargetPath = $TargetPath.TrimEnd('\')

        if ($BasePath -eq $TargetPath) {
            return "."
        }

        # Use Uri-based relative paths so this works on Windows PowerShell 5.1
        # (.NET versions that do not provide Path.GetRelativePath).
        $baseUri = New-Object System.Uri(($BasePath + "\"))
        $targetUri = New-Object System.Uri($TargetPath)
        $relativeUri = $baseUri.MakeRelativeUri($targetUri)
        $relativePath = [System.Uri]::UnescapeDataString($relativeUri.ToString())

        return $relativePath.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    }

    function Get-FullPath {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if ($Path -match '^(?:Microsoft\.PowerShell\.Core\\)?FileSystem::(.+)$') {
            $Path = $Matches[1]
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            throw "Path does not exist: $Path"
        }

        return (Resolve-Path -LiteralPath $Path).Path
    }

    function New-LogFilePath {
        param(
            [Parameter(Mandatory)]
            [string]$PreferredFolder
        )

        $logName = "PublisherConversion_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
        $candidates = @(
            (Join-Path $PreferredFolder $logName),
            (Join-Path $env:TEMP $logName)
        )

        foreach ($path in $candidates) {
            try {
                $folder = Split-Path -Parent $path

                if ($folder -and -not (Test-Path -LiteralPath $folder)) {
                    New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
                }

                # File.Create is more reliable than New-Item on some OneDrive folders.
                $stream = [System.IO.File]::Create($path)
                $stream.Close()
                return $path
            }
            catch {
                Write-Host "WARN: Could not create log file at $path"
            }
        }

        return $null
    }

    function Write-Log {
        param(
            [AllowEmptyString()]
            [string]$Message
        )

        if ($Message.Length -eq 0) {
            Write-Host ""
            if ($script:LogFile) {
                Add-Content -LiteralPath $script:LogFile -Value ""
            }
            return
        }

        $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
        Write-Host $line

        if ($script:LogFile) {
            Add-Content -LiteralPath $script:LogFile -Value $line
        }
    }

    function Get-ComFileSystemObject {
        if ($null -eq $script:ComFileSystemObject) {
            $script:ComFileSystemObject = New-Object -ComObject Scripting.FileSystemObject
        }

        return $script:ComFileSystemObject
    }

    function Get-ComSafePath {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        # Strip PowerShell provider prefixes; COM expects plain Windows paths.
        if ($Path -match '^(?:Microsoft\.PowerShell\.Core\\)?FileSystem::(.+)$') {
            $Path = $Matches[1]
        }

        if (Test-Path -LiteralPath $Path) {
            $fso = Get-ComFileSystemObject
            $item = Get-Item -LiteralPath $Path -Force

            # Use the literal input path with FSO, not Get-Item.FullName.
            # FullName expansion from 8.3 paths breaks when the profile name contains "(".
            if ($item.PSIsContainer) {
                return $fso.GetFolder($Path).ShortPath
            }

            return $fso.GetFile($Path).ShortPath
        }

        $parent = Split-Path -Parent $Path
        $leaf = Split-Path -Leaf $Path

        if ($parent -and (Test-Path -LiteralPath $parent)) {
            return Join-Path (Get-ComSafePath -Path $parent) $leaf
        }

        return $Path
    }

    function Resolve-ExistingPath {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if (-not (Test-Path -LiteralPath $Path)) {
            throw "Path does not exist: $Path"
        }

        return Get-FullPath -Path (Resolve-Path -LiteralPath $Path).Path
    }

    function Resolve-OrCreateFolder {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }

        return Get-FullPath -Path $Path
    }

    function Release-ComObject {
        param(
            [object]$ComObject
        )

        if ($null -ne $ComObject) {
            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            }
            catch {
                # Ignore release errors
            }
        }
    }

    function Invoke-ComGarbageCollection {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }

    function New-PublisherApplication {
        Write-Log "Starting Microsoft Publisher COM automation"
        $publisher = New-Object -ComObject Publisher.Application

        try {
            if ($publisher.ActiveWindow) {
                $publisher.ActiveWindow.Visible = $true
            }
        }
        catch {
            # Some hosts do not expose ActiveWindow immediately.
        }

        return $publisher
    }

    function Invoke-PublisherDoEvents {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            [System.Windows.Forms.Application]::DoEvents()
        }
        catch {
            # Optional; export still works without it.
        }
    }

    function Set-PublisherWindowVisible {
        param(
            [Parameter(Mandatory)]
            [object]$Publisher,

            [bool]$Visible = $true
        )

        try {
            if ($Publisher.ActiveWindow) {
                $Publisher.ActiveWindow.Visible = $Visible
            }
        }
        catch {
            # Ignore if no window is available.
        }
    }

    function Start-ExportWatchdog {
        param(
            [Parameter(Mandatory)]
            [string]$TempPdfPath,

            [Parameter(Mandatory)]
            [int]$TimeoutSeconds
        )

        if ($TimeoutSeconds -le 0) {
            return $null
        }

        return Start-Job -ScriptBlock {
            param($TempPdfPath, $TimeoutSeconds)

            $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

            while ((Get-Date) -lt $deadline) {
                if (Test-Path -LiteralPath $TempPdfPath) {
                    return
                }

                Start-Sleep -Seconds 2
            }

            Get-Process -Name MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force
        } -ArgumentList $TempPdfPath, $TimeoutSeconds
    }

    function Stop-ExportWatchdog {
        param(
            [System.Management.Automation.Job]$WatchdogJob
        )

        if ($null -eq $WatchdogJob) {
            return
        }

        try {
            Stop-Job -Job $WatchdogJob -ErrorAction SilentlyContinue
            Remove-Job -Job $WatchdogJob -Force -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore cleanup errors.
        }
    }

    function Invoke-PublisherPdfExport {
        param(
            [Parameter(Mandatory)]
            [object]$Document,

            [Parameter(Mandatory)]
            [string]$ComTempPdfPath
        )

        Write-Log "WORKING : Publisher is exporting (usually 30-120 seconds)"
        Write-Log "HINT    : If this stalls, check the taskbar for a Publisher window or dialog"

        Invoke-PublisherDoEvents
        Start-Sleep -Seconds 2

        $missing = [Type]::Missing
        $errors = @()

        $attempts = @(
            @{
                Name = "ExportAsFixedFormat(Format, FileName)"
                Action = { param($doc, $path) $doc.ExportAsFixedFormat($pbFixedFormatTypePDF, $path) }
            },
            @{
                Name = "ExportAsFixedFormat(Format, FileName, Intent)"
                Action = { param($doc, $path) $doc.ExportAsFixedFormat($pbFixedFormatTypePDF, $path, $pbIntentStandard) }
            },
            @{
                Name = "ExportAsFixedFormat(no accessibility tags)"
                Action = {
                    param($doc, $path)
                    $doc.ExportAsFixedFormat(
                        $pbFixedFormatTypePDF,
                        $path,
                        $pbIntentStandard,
                        $false,
                        $missing,
                        $missing,
                        $missing,
                        $missing,
                        $missing,
                        $missing,
                        $missing,
                        $missing,
                        $missing,
                        $false
                    )
                }
            }
        )

        foreach ($attempt in $attempts) {
            Write-Log "TRY     : $($attempt.Name)"
            Invoke-PublisherDoEvents

            try {
                & $attempt.Action $Document $ComTempPdfPath
                Write-Log "OK      : $($attempt.Name)"
                return
            }
            catch {
                $errors += "$($attempt.Name): $($_.Exception.Message)"
            }
        }

        throw "Publisher export failed. $($errors -join ' | ')"
    }

    function Stop-PublisherApplication {
        param(
            [object]$Publisher
        )

        if ($null -eq $Publisher) {
            return
        }

        try {
            $Publisher.Quit()
        }
        catch {
            Write-Log "WARN    : Could not quit Publisher cleanly"
            Write-Log "WARN    : $($_.Exception.Message)"
        }

        Release-ComObject $Publisher
        Release-ComObject $script:ComFileSystemObject
        $script:ComFileSystemObject = $null
        Invoke-ComGarbageCollection
    }

    function Test-ExistingPdfSkip {
        param(
            [Parameter(Mandatory)]
            [string]$PdfPath
        )

        if (-not $Skip) {
            return $false
        }

        if (-not (Test-Path -LiteralPath $PdfPath)) {
            return $false
        }

        Write-Log "SKIPPED : PDF already exists: $PdfPath"
        return $true
    }

    function Get-PublisherApplication {
        param(
            [ref]$Publisher
        )

        if ($null -eq $Publisher.Value) {
            $Publisher.Value = New-PublisherApplication
        }

        return $Publisher.Value
    }

    function Test-RequiresLocalPublisherCopy {
        param(
            [Parameter(Mandatory)]
            [string]$PubPath
        )

        if ($PubPath -match '\\OneDrive(?:\s|-)') {
            return $true
        }

        if ($PubPath -match '^\\\\') {
            return $true
        }

        $item = Get-Item -LiteralPath $PubPath -Force
        $attrs = [int]$item.Attributes

        # Offline / cloud-only placeholders (common with OneDrive).
        $offline = 0x00001000
        $recall = 0x00400000

        if (($attrs -band $offline) -ne 0) {
            return $true
        }

        if (($attrs -band $recall) -ne 0) {
            return $true
        }

        return $false
    }

    function Test-OleCompoundFile {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $header = [byte[]]::new(8)
        $stream = [System.IO.File]::OpenRead($Path)

        try {
            if ($stream.Read($header, 0, 8) -lt 8) {
                return $false
            }
        }
        finally {
            $stream.Dispose()
        }

        $ole = [byte[]]@(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
        for ($i = 0; $i -lt 8; $i++) {
            if ($header[$i] -ne $ole[$i]) {
                return $false
            }
        }

        return $true
    }

    function Get-OleCompoundKind {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if (-not (Test-OleCompoundFile -Path $Path)) {
            return "NotOle"
        }

        $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($Path))

        if ($text -match 'Microsoft Publisher|MSPublisher|PubMagic|CPublisherDoc|Quill96 Story') {
            return "Publisher"
        }

        if ($text -match 'Word\.Document|\x00WordDocument') {
            return "Word"
        }

        if ($text -match 'PowerPoint Document') {
            return "PowerPoint"
        }

        if ($text -match 'Workbook') {
            return "Excel"
        }

        return "UnknownOle"
    }

    function Get-FileMagicDescription {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $header = [byte[]]::new(8)
        $stream = [System.IO.File]::OpenRead($Path)

        try {
            if ($stream.Read($header, 0, 8) -lt 4) {
                return "too small to identify"
            }
        }
        finally {
            $stream.Dispose()
        }

        $hex = ($header | ForEach-Object { $_.ToString("X2") }) -join " "

        if ($header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0) {
            return "OLE compound document (header: $hex)"
        }

        if ($header[0] -eq 0x25 -and $header[1] -eq 0x50 -and $header[2] -eq 0x44 -and $header[3] -eq 0x46) {
            return "PDF, not Publisher (header: $hex)"
        }

        if ($header[0] -eq 0x50 -and $header[1] -eq 0x4B) {
            return "ZIP archive, not classic Publisher (header: $hex)"
        }

        if ($header[0] -eq 0x3C) {
            return "text/HTML/XML, not Publisher (header: $hex)"
        }

        return "unknown format (header: $hex)"
    }

    function Invoke-HydrateCloudFile {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )

        try {
            $buffer = [byte[]]::new(65536)
            $total = 0

            while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $total += $read
            }

            return $total
        }
        finally {
            $stream.Dispose()
        }
    }

    function New-LocalPublisherCopy {
        param(
            [Parameter(Mandatory)]
            [string]$PubPath
        )

        $tempDir = Join-Path $env:TEMP "pub2pdf"
        if (-not (Test-Path -LiteralPath $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $tempPubPath = Join-Path $tempDir ("src_{0}.pub" -f [Guid]::NewGuid().ToString("N"))

        Write-Log "STAGING : $tempPubPath"
        Write-Log "REASON  : Publisher COM needs a local copy of the source file"

        $sourceSize = (Get-Item -LiteralPath $PubPath -Force).Length
        Write-Log "SOURCEB : $sourceSize bytes before hydration"

        $hydratedBytes = Invoke-HydrateCloudFile -Path $PubPath
        Write-Log "HYDRATE : read $hydratedBytes bytes from source"

        if ($hydratedBytes -eq 0) {
            throw "Source file is empty or unreadable (may still be cloud-only): $PubPath"
        }

        Copy-Item -LiteralPath $PubPath -Destination $tempPubPath -Force

        $stagedSize = (Get-Item -LiteralPath $tempPubPath -Force).Length
        Write-Log "STAGED  : $stagedSize bytes copied to temp"

        if ($stagedSize -ne $hydratedBytes) {
            Write-Log "WARN    : staged size ($stagedSize) differs from hydrated read ($hydratedBytes)"
        }

        if (-not (Test-OleCompoundFile -Path $tempPubPath)) {
            $magic = Get-FileMagicDescription -Path $tempPubPath
            throw (
                "Staged file is not a valid Publisher/OLE document.`n" +
                "Detected: $magic`n" +
                "The .pub may be misnamed, corrupt, or not fully downloaded from OneDrive.`n" +
                "Try opening the source file manually in Publisher, or in OneDrive choose 'Always keep on this device' and rerun.`n" +
                "Source: $PubPath"
            )
        }

        return Get-ComSafePath -Path $tempPubPath
    }

    function Resolve-ComPathForLogging {
        param(
            [Parameter(Mandatory)]
            [string]$Path,

            [Parameter(Mandatory)]
            [string]$ComPath
        )

        if ($ComPath -ne $Path) {
            Write-Log "COMPATH : $ComPath"
        }
    }

    function Write-PublisherFileDiagnostics {
        param(
            [Parameter(Mandatory)]
            [string]$PubPath
        )

        try {
            $item = Get-Item -LiteralPath $PubPath -Force
            Write-Log "DIAG    : size=$($item.Length) attrs=$($item.Attributes)"

            if ($item.Length -eq 0) {
                Write-Log "HINT    : file is 0 bytes; in OneDrive choose 'Always keep on this device'"
            }
            elseif (Test-RequiresLocalPublisherCopy -PubPath $PubPath) {
                Write-Log "HINT    : source appears cloud/network-backed; script will use a local copy"
            }

            if (Test-Path -LiteralPath $PubPath) {
                Write-Log "FORMAT  : $(Get-FileMagicDescription -Path $PubPath)"
                if (-not (Test-OleCompoundFile -Path $PubPath)) {
                    Write-Log "HINT    : source bytes are not a Publisher/OLE document; Publisher will reject this file"
                }
            }
        }
        catch {
            Write-Log "DIAG    : could not inspect source file: $($_.Exception.Message)"
        }
    }

    function Open-PublisherDocument {
        param(
            [Parameter(Mandatory)]
            [object]$Publisher,

            [Parameter(Mandatory)]
            [string]$PubPath
        )

        $comPath = Get-ComSafePath -Path $PubPath
        Resolve-ComPathForLogging -Path $PubPath -ComPath $comPath

        $errors = @()

        foreach ($argumentCount in 3, 2, 1) {
            try {
                switch ($argumentCount) {
                    3 { return $Publisher.Open($comPath, $true, $false) }
                    2 { return $Publisher.Open($comPath, $true) }
                    1 { return $Publisher.Open($comPath) }
                }
            }
            catch {
                $errors += "Open($argumentCount args): $($_.Exception.Message)"
            }
        }

        throw "Publisher cannot open the file. $($errors -join ' | ')"
    }

    function Copy-PdfToDestination {
        param(
            [Parameter(Mandatory)]
            [string]$SourcePath,

            [Parameter(Mandatory)]
            [string]$DestinationPath
        )

        $destFolder = Split-Path -Parent $DestinationPath

        if ($destFolder -and -not (Test-Path -LiteralPath $destFolder)) {
            New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
        }

        if (-not (Test-Path -LiteralPath $SourcePath)) {
            throw "Temp PDF not found for copy: $SourcePath"
        }

        $bytes = [System.IO.File]::ReadAllBytes($SourcePath)
        if ($bytes.Length -eq 0) {
            throw "Temp PDF is empty: $SourcePath"
        }

        Write-Log "COPY    : $($bytes.Length) bytes to $DestinationPath"

        try {
            [System.IO.File]::WriteAllBytes($DestinationPath, $bytes)
        }
        catch {
            throw "Could not save PDF to destination (OneDrive may be blocking writes): $DestinationPath`n$($_.Exception.Message)"
        }

        if (-not (Test-Path -LiteralPath $DestinationPath)) {
            throw "PDF was not saved to destination: $DestinationPath"
        }
    }

    function Export-PublisherDocumentToPdf {
        param(
            [Parameter(Mandatory)]
            [object]$Document,

            [Parameter(Mandatory)]
            [string]$FinalPdfPath,

            [Parameter(Mandatory)]
            [object]$Publisher
        )

        $pdfFolder = Split-Path -Parent $FinalPdfPath

        if (-not (Test-Path -LiteralPath $pdfFolder)) {
            New-Item -ItemType Directory -Path $pdfFolder -Force | Out-Null
        }

        $tempDir = Join-Path $env:TEMP "pub2pdf"
        if (-not (Test-Path -LiteralPath $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $tempPdfName = "out_{0}.pdf" -f [Guid]::NewGuid().ToString("N")
        $tempPdfPath = Join-Path $tempDir $tempPdfName
        $fso = Get-ComFileSystemObject
        $comTempPdfPath = Join-Path ($fso.GetFolder($tempDir).ShortPath) $tempPdfName
        $watchdog = $null
        $copied = $false

        try {
            Write-Log "EXPORT  : $FinalPdfPath"
            Write-Log "TEMP    : $tempPdfPath"
            Resolve-ComPathForLogging -Path $tempPdfPath -ComPath $comTempPdfPath

            Set-PublisherWindowVisible -Publisher $Publisher -Visible $true
            $watchdog = Start-ExportWatchdog -TempPdfPath $tempPdfPath -TimeoutSeconds $ExportTimeoutSeconds

            Invoke-PublisherPdfExport -Document $Document -ComTempPdfPath $comTempPdfPath

            Start-Sleep -Milliseconds 500

            $actualTempPath = $null
            foreach ($candidate in @($tempPdfPath, $comTempPdfPath)) {
                if ((Test-Path -LiteralPath $candidate) -and ((Get-Item -LiteralPath $candidate).Length -gt 0)) {
                    $actualTempPath = $candidate
                    break
                }
            }

            if (-not $actualTempPath) {
                throw "Publisher export completed but temp PDF was not created: $tempPdfPath"
            }

            Copy-PdfToDestination -SourcePath $actualTempPath -DestinationPath $FinalPdfPath
            $copied = $true
        }
        catch {
            if ($ExportTimeoutSeconds -gt 0 -and -not (Test-Path -LiteralPath $tempPdfPath)) {
                throw (
                    "Publisher export did not finish within $ExportTimeoutSeconds seconds.`n" +
                    "A hidden Publisher dialog may be waiting for input - check the taskbar.`n" +
                    "Original error: $($_.Exception.Message)"
                )
            }

            throw
        }
        finally {
            Stop-ExportWatchdog -WatchdogJob $watchdog

            if ($copied) {
                foreach ($candidate in @($tempPdfPath, $comTempPdfPath)) {
                    if (Test-Path -LiteralPath $candidate) {
                        Remove-Item -LiteralPath $candidate -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }

    function Convert-OnePublisherFile {
        param(
            [Parameter(Mandatory)]
            [System.IO.FileInfo]$PubFile,

            [Parameter(Mandatory)]
            [string]$PdfPath,

            [Parameter(Mandatory)]
            [object]$Publisher
        )

        $document = $null
        $activeWindow = $null
        $tempPubPath = $null
        $pubPath = $PubFile.FullName
        $pdfPath = $PdfPath
        $openPath = $pubPath

        try {
            if (Test-RequiresLocalPublisherCopy -PubPath $pubPath) {
                $tempPubPath = New-LocalPublisherCopy -PubPath $pubPath
                $openPath = $tempPubPath
            }
            else {
                $openPath = Get-ComSafePath -Path $pubPath
            }

            Write-Log "OPENING : $openPath"
            if ($openPath -ne $pubPath) {
                Write-Log "SOURCE  : $pubPath"
            }

            if (-not (Test-OleCompoundFile -Path $openPath)) {
                throw (
                    "File is not a valid Publisher/OLE document: $(Get-FileMagicDescription -Path $openPath)`n" +
                    "Publisher reported 'This is not a Publisher file' for similar content.`n" +
                    "Source: $pubPath"
                )
            }

            try {
                $document = Open-PublisherDocument -Publisher $Publisher -PubPath $openPath
            }
            catch {
                if ($null -eq $tempPubPath) {
                    Write-Log "RETRY   : staging local copy after Publisher open failure"
                    Write-PublisherFileDiagnostics -PubPath $pubPath
                    $tempPubPath = New-LocalPublisherCopy -PubPath $pubPath
                    $openPath = $tempPubPath
                    Write-Log "OPENING : $openPath"
                    $document = Open-PublisherDocument -Publisher $Publisher -PubPath $openPath
                }
                else {
                    throw
                }
            }

            Start-Sleep -Seconds 2

            Export-PublisherDocumentToPdf -Document $document -FinalPdfPath $pdfPath -Publisher $Publisher

            Start-Sleep -Milliseconds 300

            Write-Log "SUCCESS : $pdfPath"
            return "Success"
        }
        catch {
            Write-Log "FAILED  : $pubPath"
            Write-Log "TARGET  : $pdfPath"
            Write-PublisherFileDiagnostics -PubPath $pubPath
            Write-Log "ERROR   : $($_.Exception.Message)"
            Write-Log "TYPE    : $($_.Exception.GetType().FullName)"

            if ($_.Exception.InnerException) {
                Write-Log "INNER   : $($_.Exception.InnerException.Message)"
            }

            if ($_.InvocationInfo) {
                Write-Log "AT      : $($_.InvocationInfo.PositionMessage.Trim())"
            }

            return "Failed"
        }
        finally {
            if ($null -ne $document) {
                try {
                    $document.Close()
                }
                catch {
                    Write-Log "WARN    : Could not close document cleanly: $($PubFile.FullName)"
                    Write-Log "WARN    : $($_.Exception.Message)"
                }

                Release-ComObject $document
            }

            Release-ComObject $activeWindow

            if ($tempPubPath -and (Test-Path -LiteralPath $tempPubPath)) {
                Remove-Item -LiteralPath $tempPubPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # Normalise paths
    $OutputRoot = Resolve-OrCreateFolder $OutputRoot

    if ($PSCmdlet.ParameterSetName -eq "Bulk") {
        $SourceRoot = Resolve-ExistingPath $SourceRoot
    }
    else {
        $File = Resolve-ExistingPath $File
    }

    $script:LogFile = New-LogFilePath -PreferredFolder $OutputRoot

    if (-not $script:LogFile) {
        Write-Host "WARN: Could not create a log file. Progress will show in this window only."
    }

    Write-Log "Starting conversion"
    Write-Log "Mode   : $($PSCmdlet.ParameterSetName)"
    Write-Log "Target : $OutputRoot"
    Write-Log "Log    : $script:LogFile"
    Write-Log "Skip   : $(if ($Skip) { 'existing PDFs' } else { 'no' })"
    Write-Log "Overwrite : $(if ($Overwrite) { 'yes' } else { 'no' })"

    $publisher = $null
    $success = 0
    $failed  = 0
    $skipped = 0
    $totalFiles = 0

    try {
        $openedSinceRestart = 0

        if ($PSCmdlet.ParameterSetName -eq "Single") {
            $pubFile = Get-Item -LiteralPath $File

            if ($pubFile.Extension -ne ".pub") {
                throw "Single-file mode requires a .pub file: $($pubFile.FullName)"
            }

            $totalFiles = 1

            $pdfPath = Join-Path $OutputRoot ($pubFile.BaseName + ".pdf")

            if (Test-ExistingPdfSkip -PdfPath $pdfPath) {
                $skipped++
            }
            else {
                $publisher = Get-PublisherApplication -Publisher ([ref]$publisher)

                $result = Convert-OnePublisherFile `
                    -PubFile $pubFile `
                    -PdfPath $pdfPath `
                    -Publisher $publisher

                switch ($result) {
                    "Success" { $success++ }
                    "Failed"  { $failed++ }
                    "Skipped" { $skipped++ }
                }
            }
        }
        else {
            Write-Log "Source : $SourceRoot"

            $files = Get-ChildItem `
                -LiteralPath $SourceRoot `
                -Filter "*.pub" `
                -Recurse `
                -File

            $totalFiles = $files.Count

            Write-Log "Found  : $totalFiles Publisher files"
            if ($RestartEvery -gt 0) {
                Write-Log "Restart : every $RestartEvery opened files"
            }

            foreach ($pubFile in $files) {
                try {
                    $relativeFolder = Get-RelativePathCompat `
                        -BasePath $SourceRoot `
                        -TargetPath $pubFile.DirectoryName

                    if ($relativeFolder -eq ".") {
                        $targetFolder = $OutputRoot
                    }
                    else {
                        $targetFolder = Join-Path $OutputRoot $relativeFolder
                    }

                    $pdfPath = Join-Path $targetFolder ($pubFile.BaseName + ".pdf")

                    if (Test-ExistingPdfSkip -PdfPath $pdfPath) {
                        $skipped++
                        continue
                    }

                    $publisher = Get-PublisherApplication -Publisher ([ref]$publisher)

                    $result = Convert-OnePublisherFile `
                        -PubFile $pubFile `
                        -PdfPath $pdfPath `
                        -Publisher $publisher

                    switch ($result) {
                        "Success" {
                            $success++
                            $openedSinceRestart++
                        }
                        "Failed" {
                            $failed++
                            $openedSinceRestart++
                        }
                        "Skipped" { $skipped++ }
                    }

                    if ($RestartEvery -gt 0 -and $openedSinceRestart -ge $RestartEvery) {
                        Write-Log "RESTART : Recycling Publisher after $openedSinceRestart opened files"
                        Stop-PublisherApplication -Publisher $publisher
                        $publisher = New-PublisherApplication
                        $openedSinceRestart = 0
                    }
                }
                catch {
                    Write-Log "FAILED  : $($pubFile.FullName)"
                    Write-Log "ERROR   : $($_.Exception.Message)"
                    Write-Log "TYPE    : $($_.Exception.GetType().FullName)"
                    $failed++
                }
            }
        }
    }
    finally {
        Stop-PublisherApplication -Publisher $publisher

        Write-Log ""
        Write-Log "============================="
        Write-Log "Conversion complete"
        Write-Log "Files   : $totalFiles"
        Write-Log "Success : $success"
        Write-Log "Failed  : $failed"
        Write-Log "Skipped : $skipped"
        Write-Log "Log     : $script:LogFile"
        Write-Log "============================="
    }
