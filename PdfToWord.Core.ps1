<#
    PdfToWord.Core.ps1 - shared implementation used by Convert-PdfToWord.ps1 and
    PdfToWordWorker.ps1. Dot-source this file; it defines functions and constants only.

    Nothing in here writes to stdout. The worker relies on that: its stdout is a JSON
    channel and any stray output would corrupt the protocol.
#>

Set-StrictMode -Version Latest

# WdSaveFormat values. WebOptions/Encoding say which mechanism makes the output UTF-8:
# for HTML the Encoding argument of SaveAs2 is silently ignored - the document's
# WebOptions.Encoding is what actually takes effect.
$script:PdfToWordFormats = @{
    docx     = @{ Extension = '.docx'; WdFormat = 12; WebOptions = $false; Encoding = $false }
    doc      = @{ Extension = '.doc';  WdFormat = 0;  WebOptions = $false; Encoding = $false }
    rtf      = @{ Extension = '.rtf';  WdFormat = 6;  WebOptions = $false; Encoding = $false }
    txt      = @{ Extension = '.txt';  WdFormat = 2;  WebOptions = $false; Encoding = $true  }  # wdFormatText
    odt      = @{ Extension = '.odt';  WdFormat = 23; WebOptions = $false; Encoding = $false }
    html     = @{ Extension = '.html'; WdFormat = 10; WebOptions = $true;  Encoding = $false }  # wdFormatFilteredHTML
    htmlfull = @{ Extension = '.html'; WdFormat = 8;  WebOptions = $true;  Encoding = $false }  # wdFormatHTML
    pdf      = @{ Extension = '.pdf';  WdFormat = 17; WebOptions = $false; Encoding = $false }
}

$script:PdfToWordMsoEncodingUtf8              = 65001
$script:PdfToWordMsoAutomationSecurityDisable = 3
$script:PdfToWordWdDoNotSaveChanges           = 0
$script:PdfToWordWdStatisticPages             = 2
$script:PdfToWordMissing                      = [Type]::Missing
$script:PdfToWordDefaultImagesFolder          = 'source_folder'

function Resolve-TargetPath {
    param([string]$Path, [string]$Base)
    if (-not [IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path -Path $Base -ChildPath $Path
    }
    return [IO.Path]::GetFullPath($Path)
}

<#
    Converts every non-PNG image in an HTML export folder to PNG, rewrites the matching
    references in the .html, then deletes the originals. Word's WebOptions.AllowPNG
    already does this for most pictures; this sweep catches whatever it leaves behind
    (CMYK JPEGs, linked pictures, unusual profiles). Non-image companions such as
    filelist.xml and themedata.thmx are left alone.
#>
function Convert-ExportImagesToPng {
    param([string]$Folder, [string]$HtmlPath)

    $imageExtensions = @('.jpg', '.jpeg', '.jpe', '.gif', '.bmp', '.tif', '.tiff', '.emf', '.wmf', '.ico')
    $replacements = [ordered]@{}

    try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop }
    catch { Write-Warning "Cannot load System.Drawing; images were left in their original format."; return }

    foreach ($file in @(Get-ChildItem -LiteralPath $Folder -File -Force -ErrorAction SilentlyContinue)) {
        if ($imageExtensions -notcontains $file.Extension.ToLowerInvariant()) { continue }

        $baseName = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        $targetName = "$baseName.png"
        $i = 0
        while ((Test-Path -LiteralPath (Join-Path $Folder $targetName)) -or
               @($replacements.Values) -contains $targetName) {
            $i++
            $targetName = "$baseName-$i.png"
        }
        $targetPath = Join-Path $Folder $targetName

        try {
            $image = [System.Drawing.Image]::FromFile($file.FullName)
            try {
                if ($file.Extension -match '^\.(emf|wmf)$') {
                    # Vector: rasterise at its native size (typically 96 dpi).
                    $w = [int][Math]::Max(1, [Math]::Ceiling($image.Width))
                    $h = [int][Math]::Max(1, [Math]::Ceiling($image.Height))
                    $bitmap = New-Object System.Drawing.Bitmap($w, $h)
                    $bitmap.SetResolution(96, 96)
                    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                    try {
                        $graphics.Clear([System.Drawing.Color]::Transparent)
                        $graphics.DrawImage($image, 0, 0, $w, $h)
                    }
                    finally { $graphics.Dispose() }
                    try { $bitmap.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png) }
                    finally { $bitmap.Dispose() }
                }
                else {
                    $image.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png)
                }
            }
            finally { $image.Dispose() }

            $replacements[$file.Name] = $targetName
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            Write-Verbose "Converted image '$($file.Name)' -> '$targetName'"
        }
        catch {
            if (Test-Path -LiteralPath $targetPath) {
                Remove-Item -LiteralPath $targetPath -Force -ErrorAction SilentlyContinue
            }
            Write-Warning "Could not convert '$($file.Name)' to PNG: $($_.Exception.Message)"
        }
    }

    if ($replacements.Count -eq 0) { return }

    # Fix the links. File names are unique enough that replacing the bare name is safe,
    # but the folder-qualified form is replaced first for good measure.
    $html = [IO.File]::ReadAllText($HtmlPath, [Text.Encoding]::UTF8)
    foreach ($old in $replacements.Keys) {
        $html = $html.Replace($old, $replacements[$old])
    }
    [IO.File]::WriteAllText($HtmlPath, $html, (New-Object Text.UTF8Encoding($false)))
}

# ---------- watchdog ----------
# A Word modal dialog would otherwise block a conversion forever with no way out. This
# has to be C#: a PowerShell script block cannot run on a timer thread (no runspace).
function Initialize-PdfToWordWatchdog {
    if ('PdfToWord.Watchdog' -as [type]) { return }
    Add-Type -TypeDefinition @'
namespace PdfToWord {
    public static class Watchdog {
        public static System.Threading.Timer Timer;
        public static volatile bool TimedOut;
        public static void Start(int pid, int milliseconds) {
            TimedOut = false;
            Timer = new System.Threading.Timer(_ => {
                TimedOut = true;
                try { System.Diagnostics.Process.GetProcessById(pid).Kill(); } catch { }
            }, null, milliseconds, System.Threading.Timeout.Infinite);
        }
        public static void Stop() {
            if (Timer != null) { Timer.Dispose(); Timer = null; }
        }
        public static void Reset() { TimedOut = false; }
    }
}
'@
}

function Test-PdfToWordWatchdogTripped {
    if (-not ('PdfToWord.Watchdog' -as [type])) { return $false }
    return [bool][PdfToWord.Watchdog]::TimedOut
}

# ---------- Word sessions ----------
<#
    Starts (or optionally attaches to) a Word instance and records everything needed to
    shut it down safely.

    CreateInstance is the only way the worker ever gets Word. It spawns a fresh
    WINWORD.EXE /Automation -Embedding process; shell launches attach to an existing
    one, which is exactly why the worker must never use Start-Process. The process id is
    found with a before/after diff, and a session-wide named mutex serialises that diff
    so several workers can start in parallel without claiming each other's Word.

    If CreateInstance ever does not produce a new process we deliberately mark the
    session as NOT owned. That can leak a hidden Word, but it can never quit or kill a
    Word we did not create - the wrong trade only in the other direction.
#>
function New-WordSession {
    [CmdletBinding()]
    param(
        [switch]$AllowAttach,
        [switch]$Visible,
        [switch]$NoRegistryFix
    )

    $session = [pscustomobject]@{
        Word             = $null
        Pid              = $null
        Owned            = $false
        PreviousVisible  = $null
        PreviousAlerts   = $null
        PreviousSecurity = $null
        PdfWarningSet    = $false
        Version          = $null
    }

    if ($AllowAttach) {
        try {
            $session.Word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            Write-Verbose 'Attached to the already-running Word instance; it will not be quit.'
        }
        catch { }
    }

    if (-not $session.Word) {
        $mutex = New-Object System.Threading.Mutex($false, 'Local\PdfToWord.WordActivation')
        $acquired = $false
        try {
            try { $acquired = $mutex.WaitOne(30000) }
            catch [System.Threading.AbandonedMutexException] { $acquired = $true }
            if (-not $acquired) {
                throw [System.Exception]::new('Timed out waiting for the Word activation lock.')
            }

            $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $session.Word = [Activator]::CreateInstance([type]::GetTypeFromProgID('Word.Application'))
            $after = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $session.Pid = @($after | Where-Object { $before -notcontains $_ }) | Select-Object -First 1
        }
        finally {
            if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
            try { $mutex.Dispose() } catch { }
        }

        if ($session.Pid) {
            $session.Owned = $true
        }
        else {
            Write-Warning 'Word started but its process id could not be identified; it will not be quit or killed by this session.'
        }
    }

    if ($session.Owned) {
        try { $session.Word.Visible = [bool]$Visible } catch { }
    }
    else {
        $session.PreviousVisible = $session.Word.Visible
        $session.PreviousAlerts  = $session.Word.DisplayAlerts
        try { $session.PreviousSecurity = $session.Word.AutomationSecurity }
        catch { $session.PreviousSecurity = $null }
    }

    try { $session.Word.DisplayAlerts = 0 } catch { }   # wdAlertsNone
    try { $session.Word.AutomationSecurity = $script:PdfToWordMsoAutomationSecurityDisable } catch { }
    try { $session.Version = [string]$session.Word.Version } catch { }

    # Suppress Word's modal "convert this PDF?" prompt, which stalls an invisible Word
    # instance forever.
    if (-not $NoRegistryFix) {
        try {
            $regPath = "HKCU:\Software\Microsoft\Office\$($session.Version)\Word\Options"
            if (-not (Test-Path -LiteralPath $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            $optionKey = Get-ItemProperty -LiteralPath $regPath -ErrorAction SilentlyContinue
            $current = $null
            if ($optionKey -and $optionKey.PSObject.Properties['DisableConvertPdfWarning']) {
                $current = $optionKey.DisableConvertPdfWarning
            }
            if ($current -ne 1) {
                New-ItemProperty -LiteralPath $regPath -Name 'DisableConvertPdfWarning' -Value 1 -PropertyType DWord -Force | Out-Null
                $session.PdfWarningSet = $true
            }
        }
        catch {
            throw [System.Exception]::new("Could not suppress Word's PDF conversion prompt (HKCU:\Software\Microsoft\Office\$($session.Version)\Word\Options\DisableConvertPdfWarning): $($_.Exception.Message)")
        }
    }

    return $session
}

<#
    Releases the COM object and, when the session owns Word, quits it. If Quit is
    ignored the owned process is terminated - never any other WINWORD.EXE. Pass
    -KeepVisible for the CLI's -KeepOpen, which hands Word over to the user.
#>
function Close-WordSession {
    [CmdletBinding()]
    param($Session, [switch]$KeepVisible)

    if (-not $Session -or -not $Session.Word) { return }

    $word   = $Session.Word
    $ownPid = $Session.Pid
    $owned  = $Session.Owned

    $processAlive = $true
    if ($ownPid) {
        $processAlive = [bool](Get-Process -Id $ownPid -ErrorAction SilentlyContinue)
    }

    if ($processAlive) {
        if ($owned) {
            if ($KeepVisible) {
                try { $word.Visible = $true } catch { }
            }
            else {
                try { $word.Quit($script:PdfToWordWdDoNotSaveChanges) } catch { }
            }
        }
        else {
            # Hand an attached Word back the way we found it.
            if ($null -ne $Session.PreviousAlerts)   { try { $word.DisplayAlerts = $Session.PreviousAlerts } catch { } }
            if ($null -ne $Session.PreviousVisible)  { try { $word.Visible = $Session.PreviousVisible } catch { } }
            if ($null -ne $Session.PreviousSecurity) { try { $word.AutomationSecurity = $Session.PreviousSecurity } catch { } }
        }
    }

    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    $Session.Word = $null

    if ($owned -and -not $KeepVisible -and $ownPid) {
        $alive = $true
        for ($i = 0; $i -lt 10; $i++) {
            if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { $alive = $false; break }
            Start-Sleep -Milliseconds 400
        }
        if ($alive) {
            Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue
            Write-Verbose "Terminated orphaned Word process $ownPid."
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ---------- one document through an existing Word session ----------
<#
    Opens the PDF in the supplied Word instance, saves it, and closes the document.
    Every COM object is released in the finally block. The watchdog is armed only when
    a pid is supplied, so an attached user Word is never killed.
#>
function Invoke-WordDocumentConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Word,
        [Parameter(Mandatory)][hashtable]$Spec,
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][string]$OutFile,
        [switch]$AllowPng,
        [int]$WatchdogSeconds = 0,
        [int]$WatchdogPid = 0,
        [switch]$LeaveDocumentOpen
    )

    $doc   = $null
    $pages = 0
    $armed = $false

    try {
        if ($WatchdogSeconds -gt 0 -and $WatchdogPid -gt 0) {
            Initialize-PdfToWordWatchdog
            [PdfToWord.Watchdog]::Start([int]$WatchdogPid, [int]($WatchdogSeconds * 1000))
            $armed = $true
        }

        Write-Verbose "Opening '$PdfPath' in Word"
        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        $doc = $Word.Documents.Open($PdfPath, $false, $false, $false)

        try { $pages = [int]$doc.ComputeStatistics($script:PdfToWordWdStatisticPages) } catch { }

        if ($Spec.WebOptions) {
            try { $doc.WebOptions.Encoding = $script:PdfToWordMsoEncodingUtf8 } catch { }
            # Ask Word to export pictures as PNG rather than JPEG.
            if ($AllowPng) { try { $doc.WebOptions.AllowPNG = $true } catch { } }
        }

        Write-Verbose "Saving as -> '$OutFile'"
        if ($Spec.Encoding) {
            # SaveAs2(FileName, FileFormat, LockComments, Password, AddToRecentFiles,
            #         WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts,
            #         SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding)
            $doc.SaveAs2($OutFile, $Spec.WdFormat, $false, $script:PdfToWordMissing, $false,
                         $script:PdfToWordMissing, $false, $false, $false, $false, $false,
                         $script:PdfToWordMsoEncodingUtf8)
        }
        else {
            $doc.SaveAs2($OutFile, $Spec.WdFormat)
        }

        if ($armed) { [PdfToWord.Watchdog]::Stop() }

        if (-not (Test-Path -LiteralPath $OutFile)) {
            throw "Word reported success but '$OutFile' does not exist."
        }
        if ($armed -and (Test-PdfToWordWatchdogTripped)) {
            throw [System.TimeoutException]::new("Word was terminated by the watchdog after $WatchdogSeconds s.")
        }

        return [pscustomobject]@{ Pages = $pages }
    }
    catch {
        if ($armed -and (Test-PdfToWordWatchdogTripped)) {
            throw [System.TimeoutException]::new("Word did not finish within $WatchdogSeconds s and was terminated. $($_.Exception.Message)")
        }
        throw
    }
    finally {
        if ('PdfToWord.Watchdog' -as [type]) {
            try { [PdfToWord.Watchdog]::Stop() } catch { }
        }

        if ($doc) {
            if (-not $LeaveDocumentOpen) {
                try { $doc.Close($script:PdfToWordWdDoNotSaveChanges) } catch { }
            }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
            $doc = $null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ---------- validation and destination resolution ----------
<#
    Shared by the CLI and the worker. Throws ArgumentException for anything the caller
    can fix (missing file, not a PDF, folder-shaped output, existing output without
    -Force). Creates the output directory as a side effect, exactly like the original
    script did.
#>
function Resolve-PdfToWordJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$OutPath,
        [Parameter(Mandatory)][string]$Format,
        [switch]$Force,
        [switch]$Unblock,
        [string]$BasePath = $PWD.ProviderPath
    )

    if (-not $script:PdfToWordFormats.ContainsKey($Format)) {
        throw [System.ArgumentException]::new("Unknown format '$Format'. Supported formats: $($script:PdfToWordFormats.Keys -join ', ').")
    }

    # ---------- validate input ----------
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw [System.ArgumentException]::new("Source file not found: $Path")
    }

    $pdf = (Resolve-Path -LiteralPath $Path).ProviderPath

    if ([IO.Path]::GetExtension($pdf) -ne '.pdf') {
        throw [System.ArgumentException]::new("$pdf is not a PDF. This script converts PDF input only.")
    }

    if ($Unblock) {
        try { Unblock-File -LiteralPath $pdf -ErrorAction Stop }
        catch { Write-Warning "Could not unblock the PDF: $($_.Exception.Message)" }
    }

    # ---------- work out the destination ----------
    $spec = $script:PdfToWordFormats[$Format]
    $defaultName = [IO.Path]::GetFileNameWithoutExtension($pdf) + $spec.Extension

    if ($OutPath) {
        $trimmed = $OutPath.TrimEnd('\', '/')
        if (-not $trimmed) {
            throw [System.ArgumentException]::new("OutPath '$OutPath' does not name a file.")
        }

        $outFull = Resolve-TargetPath -Path $trimmed -Base $BasePath

        # Word hangs when handed a folder-shaped path, so refuse folders up front with
        # an actionable message rather than risking a frozen converter.
        if ($OutPath -match '[\\/]+$' -or (Test-Path -LiteralPath $outFull -PathType Container)) {
            $suggested = Join-Path $trimmed $defaultName
            throw [System.ArgumentException]::new("OutPath '$OutPath' is a folder. Pass the destination file instead: -OutPath '$suggested'.")
        }

        if (-not [IO.Path]::GetExtension($outFull)) {
            $outFull += $spec.Extension
        }
        elseif ([IO.Path]::GetExtension($outFull) -ne $spec.Extension) {
            Write-Warning "Output extension '$([IO.Path]::GetExtension($outFull))' does not match -Format $Format ('$($spec.Extension)')."
        }
    }
    else {
        $outFull = Join-Path (Split-Path -Parent $pdf) $defaultName
    }

    $outName = [IO.Path]::GetFileName($outFull)
    if ($outName -match '^\.' -or $outName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw [System.ArgumentException]::new("Invalid output file name '$outName' (from '$outFull').")
    }

    if (($outFull -eq $pdf) -and ($spec.WdFormat -eq 17)) {
        throw [System.ArgumentException]::new("Refusing to overwrite the source PDF. Choose a different -OutPath.")
    }

    if ((Test-Path -LiteralPath $outFull) -and -not $Force) {
        throw [System.ArgumentException]::new("Output already exists: $outFull - use -Force to overwrite.")
    }
    $outExisted = Test-Path -LiteralPath $outFull

    $outDir = Split-Path -Parent $outFull
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    return [pscustomobject]@{
        Pdf        = $pdf
        OutFile    = $outFull
        OutDir     = $outDir
        OutExisted = $outExisted
        Spec       = $spec
        Format     = $Format
    }
}

# ---------- HTML post-processing ----------
<#
    Renames Word's "<stem>_files" folder to ImagesFolderName, repairs the links, and
    (for ImageFormat png) converts every non-PNG image in that folder. Safe to call for
    non-HTML formats: it returns immediately when the format has no web options.
#>
function Complete-PdfToWordOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Spec,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ImagesFolderName = $script:PdfToWordDefaultImagesFolder,
        [ValidateSet('png', 'original')][string]$ImageFormat = 'png',
        [datetime]$Since = [datetime]::MinValue,
        [switch]$Force
    )

    if (-not $Spec.WebOptions) { return }

    if ($ImagesFolderName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
        $ImagesFolderName -match '^\.+$' -or $ImagesFolderName -match '[\\/]') {
        throw [System.ArgumentException]::new("Invalid -ImagesFolderName '$ImagesFolderName'.")
    }

    $outDir       = Split-Path -Parent $OutputPath
    $stem         = [IO.Path]::GetFileNameWithoutExtension($OutputPath)
    $targetFolder = Join-Path $outDir $ImagesFolderName

    # Word's HTML companion folder is localized: English "<stem>_files", French
    # "<stem>_fichiers", German "<stem>_Dateien", ... It is also URL-encoded in the
    # HTML when the name has spaces. Find it by what the HTML references, then by
    # the "<stem>_" prefix, rather than assuming an English suffix.
    $htmlText = $null
    if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
        $htmlText = [IO.File]::ReadAllText($OutputPath, [Text.Encoding]::UTF8)
    }

    $wordFolder = $null
    $wordFolderName = $null
    if (Test-Path -LiteralPath $outDir -PathType Container) {
        $notBefore = if ($Since -eq [datetime]::MinValue) { $Since } else { $Since.AddSeconds(-2) }
        $candidateDirs = @(Get-ChildItem -LiteralPath $outDir -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -ne $ImagesFolderName -and
                $_.LastWriteTime -ge $notBefore
            })

        $referencedNames = @()
        if ($htmlText) {
            foreach ($m in [regex]::Matches($htmlText, '(?i)(?:src|href)\s*=\s*["'']?([^"''\s>]+)')) {
                $url = $m.Groups[1].Value
                $segment = ($url -split '[\\/]')[0]
                if ($segment) {
                    try { $referencedNames += [Uri]::UnescapeDataString($segment) }
                    catch { $referencedNames += $segment }
                }
            }
        }

        if ($referencedNames.Count -gt 0) {
            $wordFolder = $candidateDirs |
                Where-Object { $referencedNames -contains $_.Name } |
                Select-Object -First 1
        }

        if (-not $wordFolder) {
            $prefixed = @($candidateDirs | Where-Object {
                $_.Name.StartsWith($stem + '_', [StringComparison]::OrdinalIgnoreCase)
            })
            if ($prefixed.Count -eq 1) {
                $wordFolder = $prefixed[0]
            }
            elseif ($prefixed.Count -gt 1) {
                $wordFolder = $prefixed | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            }
        }

        if ($wordFolder) { $wordFolderName = $wordFolder.Name }
    }

    if ($wordFolder -and
        ([IO.Path]::GetFullPath($wordFolder.FullName) -ne [IO.Path]::GetFullPath($targetFolder))) {

        if (Test-Path -LiteralPath $targetFolder) {
            if ($Force) {
                Remove-Item -LiteralPath $targetFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                throw [System.ArgumentException]::new(
                    "Image folder already exists: $targetFolder - use -Force to overwrite.")
            }
        }

        Move-Item -LiteralPath $wordFolder.FullName -Destination $targetFolder -ErrorAction Stop

        # Patch every form the encoder may have produced for the actual (localized)
        # folder name, not just the English "<stem>_files".
        if ($htmlText) {
            $patched = $htmlText
            $folderCandidates = @(
                $wordFolderName
                [Uri]::EscapeDataString($wordFolderName)
                [Uri]::EscapeUriString($wordFolderName)
            ) | Select-Object -Unique
            foreach ($candidate in $folderCandidates) {
                if ($candidate -and $candidate -ne $ImagesFolderName) {
                    $patched = $patched.Replace($candidate, $ImagesFolderName)
                }
            }
            if ($patched -ne $htmlText) {
                [IO.File]::WriteAllText($OutputPath, $patched, (New-Object Text.UTF8Encoding($false)))
            }
        }
    }

    # ---------- force PNG images ----------
    # AllowPNG handles most pictures; this sweep guarantees the rest (odd JPEGs, vector
    # metafiles, linked images) are PNG too, and repairs the links.
    if ($ImageFormat -eq 'png' -and (Test-Path -LiteralPath $targetFolder -PathType Container)) {
        Convert-ExportImagesToPng -Folder $targetFolder -HtmlPath $OutputPath
        $pngCount = @(Get-ChildItem -LiteralPath $targetFolder -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -eq '.png' }).Count
        Write-Verbose "Image folder '$ImagesFolderName' holds $pngCount PNG file(s)."
    }
}

<#
    Names of the extra files/folders this run created beside the output (Word's HTML
    image folder, and anything else sharing the output stem). Only items written since
    $Since are reported, so a pre-existing file is never blamed on this conversion.
#>
function Get-PdfToWordCompanions {
    param(
        [Parameter(Mandatory)][hashtable]$Spec,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$SourcePath,
        [string]$ImagesFolderName = $script:PdfToWordDefaultImagesFolder,
        [Parameter(Mandatory)][datetime]$Since
    )

    $outDir = Split-Path -Parent $OutputPath
    $stem   = [IO.Path]::GetFileNameWithoutExtension($OutputPath)

    return @(Get-ChildItem -LiteralPath $outDir -Force -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.FullName -ne $OutputPath) -and
            ($_.FullName -ne $SourcePath) -and
            ($_.LastWriteTime -ge $Since) -and
            (($_.Name -like "$stem*") -or ($Spec.WebOptions -and $_.Name -eq $ImagesFolderName))
        } |
        Select-Object -ExpandProperty Name)
}

# ---------- one-shot conversion (CLI) ----------
<#
    The complete operation behind Convert-PdfToWord.ps1: validate, convert with up to
    $Attempts tries, post-process, report. Each attempt gets its own Word session; a
    retry never reuses an instance that just failed.
#>
function Invoke-PdfToWordConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$OutPath,
        [Parameter(Mandatory)][string]$Format = 'docx',
        [string]$ImagesFolderName = $script:PdfToWordDefaultImagesFolder,
        [ValidateSet('png', 'original')][string]$ImageFormat = 'png',
        [switch]$Force,
        [switch]$ShowWord,
        [switch]$KeepOpen,
        [int]$Attempts = 2,
        [int]$TimeoutSeconds = 120,
        [switch]$SkipPdfWarningFix,
        [switch]$Unblock
    )

    if ($Attempts -lt 1 -or $Attempts -gt 5) {
        throw [System.ArgumentException]::new('-Attempts must be between 1 and 5.')
    }
    if ($TimeoutSeconds -lt 0 -or $TimeoutSeconds -gt 86400) {
        throw [System.ArgumentException]::new('-TimeoutSeconds must be between 0 and 86400.')
    }

    $started = Get-Date
    $job = Resolve-PdfToWordJob -Path $Path -OutPath $OutPath -Format $Format -Force:$Force -Unblock:$Unblock

    $attempt = 0
    $info = $null
    $pdfWarningSet = $false

    while ($true) {
        $attempt++
        $session = $null

        try {
            $session = New-WordSession -AllowAttach -Visible:$ShowWord -NoRegistryFix:$SkipPdfWarningFix
            $pdfWarningSet = $pdfWarningSet -or [bool]$session.PdfWarningSet

            $watchdogPid = 0
            if ($session.Owned -and $session.Pid -and -not $KeepOpen) {
                $watchdogPid = [int]$session.Pid
            }

            $info = Invoke-WordDocumentConversion -Word $session.Word -Spec $job.Spec -PdfPath $job.Pdf `
                        -OutFile $job.OutFile -AllowPng:($ImageFormat -eq 'png') `
                        -WatchdogSeconds $(if ($KeepOpen) { 0 } else { $TimeoutSeconds }) `
                        -WatchdogPid $watchdogPid -LeaveDocumentOpen:$KeepOpen
            break
        }
        catch {
            $message = $_.Exception.Message

            # Clean up anything a failed attempt left half-written; a truncated
            # document that looks like a success is worse than none at all.
            if (-not $job.OutExisted -and (Test-Path -LiteralPath $job.OutFile)) {
                Remove-Item -LiteralPath $job.OutFile -Force -ErrorAction SilentlyContinue
            }

            if ($attempt -ge $Attempts) {
                if ($_.Exception -is [System.TimeoutException]) {
                    throw [System.TimeoutException]::new(
                        "PDF conversion timed out after $TimeoutSeconds s for '$Path' -> '$($job.OutFile)'. $message")
                }
                throw [System.Exception]::new(
                    "PDF conversion failed for '$Path' -> '$($job.OutFile)'. $message")
            }

            $delay = 3
            if ($_.Exception -is [System.TimeoutException]) {
                # The watchdog killed Word a moment ago; force-killing repeatedly leaves
                # the COM server unable to start cleanly, so give it time to settle.
                $delay = 10
            }
            Write-Warning "Attempt $attempt of $Attempts failed: $message Retrying with a fresh Word instance."
            Start-Sleep -Seconds $delay
        }
        finally {
            if ($session) {
                if ($KeepOpen -and $session.Owned) {
                    Close-WordSession -Session $session -KeepVisible
                }
                else {
                    Close-WordSession -Session $session
                }
            }
        }
    }

    Complete-PdfToWordOutput -Spec $job.Spec -OutputPath $job.OutFile `
        -ImagesFolderName $ImagesFolderName -ImageFormat $ImageFormat -Since $started -Force:$Force

    $companions = Get-PdfToWordCompanions -Spec $job.Spec -OutputPath $job.OutFile `
        -SourcePath $job.Pdf -ImagesFolderName $ImagesFolderName -Since $started

    return [pscustomobject]@{
        Source         = $job.Pdf
        Output         = $job.OutFile
        Format         = $Format
        Pages          = $info.Pages
        SizeKB         = [math]::Round((Get-Item -LiteralPath $job.OutFile).Length / 1KB, 1)
        Seconds        = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
        Attempts       = $attempt
        CompanionFiles = @($companions)
        PdfWarningSet  = [bool]$pdfWarningSet
    }
}
