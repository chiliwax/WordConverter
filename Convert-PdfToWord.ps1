<#
.SYNOPSIS
    Converts a PDF into an editable document using Microsoft Word's PDF reflow engine.

.DESCRIPTION
    Drives Microsoft Word through COM: opens the PDF, lets Word convert it into a real
    editable document, then saves it in the format you ask for.

    Word performs NO OCR. A scanned (image-only) PDF converts into a document full of
    pictures with no editable text. Layout is a reflow, not a facsimile: columns, floats
    and complex tables shift. Encrypted PDFs fail with an error.

    Every attempt runs under a watchdog (see -TimeoutSeconds) and is retried with a
    fresh Word instance if it fails or stalls, because a leftover Word process in a bad
    state otherwise hangs automation indefinitely.

    The source PDF is never modified or overwritten.

.PARAMETER Path
    The PDF to convert. Relative paths are resolved against the current directory.

.PARAMETER OutPath
    Destination FILE, not a folder. Defaults to the PDF's own folder, same base name,
    new extension (.\report.pdf -> .\report.docx).

    Pass a folder path and you get an error telling you what to use instead - Word
    hangs when handed a folder-shaped path, so this script refuses those outright.

.PARAMETER Format
    Output format. Default is docx.

        docx      editable Word document (default)
        doc       legacy Word 97-2003
        rtf       Rich Text Format
        txt       plain text, UTF-8 (with BOM)
        odt       OpenDocument Text
        html      Web Page, FILTERED (wdFormatFilteredHTML) - clean HTML with minimal
                  Word styling. Any raster image Word keeps from the PDF is written
                  to a <ImagesFolderName>\ folder beside the .html (see
                  -ImagesFolderName), so keep the two together. Documents without
                  images give a single .html file.
        htmlfull  Web Page, full (wdFormatHTML) - Word's XML-laden HTML; also puts
                  images in a <ImagesFolderName>\ folder
        pdf       PDF, re-flowed through Word

    html, htmlfull and txt are written as UTF-8. Without that, Word falls back to the
    system ANSI codepage and mangles accented characters.

.PARAMETER Force
    Overwrite OutPath if it already exists. Without this, an existing file is an error.

.PARAMETER ShowWord
    Leave Word visible during the conversion (useful for watching/debugging it).
    By default Word runs hidden.

.PARAMETER KeepOpen
    Do not close the converted document when finished, and leave Word visible.
    Also disables the watchdog, since the point is to hand Word over to you.

.PARAMETER Attempts
    How many times to try before giving up. Default 2. Each retry starts a fresh Word
    instance and waits a few seconds first.

.PARAMETER TimeoutSeconds
    Per-attempt watchdog. If an attempt has not finished after this many seconds, the
    Word process this script started is terminated and the attempt counts as failed
    (and may be retried). Default 120. Use 0 to disable.

.PARAMETER SkipPdfWarningFix
    Do not touch the registry. By default the script sets the per-user DWORD
    HKCU\Software\Microsoft\Office\<ver>\Word\Options\DisableConvertPdfWarning = 1,
    which suppresses Word's modal "Word will now convert your PDF..." dialog. Without
    it, that dialog blocks an invisible Word instance and the script hangs.

.PARAMETER Unblock
    Strip the downloaded-file marker (Zone.Identifier) from the PDF first. Files from
    the internet can open in Word's Protected View, which interferes with automation.

.PARAMETER ImagesFolderName
    Name of the folder Word's HTML export uses for images. Word itself always names it
    "<output name>_files"; the script renames it to this value after saving and
    rewrites the links in the .html to match. Default: source_folder.

    Only used by -Format html and htmlfull. Converting several documents into the same
    folder with the default name would make them share (and overwrite) one image
    folder, so pass a distinct name per document in that case.

.PARAMETER ImageFormat
    html and htmlfull only. Default png: every image Word exports is PNG (Word is asked
    for PNG natively, and anything it leaves in another format is converted afterwards,
    with the HTML links repaired). Use original to keep Word's own choice, which is
    usually JPEG.

.EXAMPLE
    .\Convert-PdfToWord.ps1 -Path .\sample-with-table.pdf
    Creates sample-with-table.docx next to the PDF.

.EXAMPLE
    .\Convert-PdfToWord.ps1 -Path .\sample-with-table.pdf -Format html -Force
    Creates filtered, UTF-8 HTML (plus a <name>_files folder if the document has
    images that Word exports separately).

.EXAMPLE
    .\Convert-PdfToWord.ps1 -Path .\report.pdf -Format rtf -OutPath .\out\report.rtf
    Writes to .\out\report.rtf, creating .\out\ if needed.

.EXAMPLE
    Get-ChildItem *.pdf | ForEach-Object { .\Convert-PdfToWord.ps1 -Path $_.FullName -Format docx }
    Batch-convert every PDF in the current folder.

.EXAMPLE
    .\Convert-PdfToWord.ps1 -Path .\scan.pdf -ShowWord
    Watch the conversion happen in a visible Word window.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [Alias('FullName', 'LiteralPath')]
    [string]$Path,

    [string]$OutPath,

    [ValidateSet('docx', 'doc', 'rtf', 'txt', 'odt', 'html', 'htmlfull', 'pdf')]
    [string]$Format = 'docx',

    # Word names its image folder "<output stem>_files". The script renames it to this
    # afterwards and repairs the links in the HTML, so the page keeps working.
    [ValidateNotNullOrEmpty()]
    [string]$ImagesFolderName = 'source_folder',

    # html/htmlfull only: every image Word exports ends up as a .png.
    [ValidateSet('png', 'original')]
    [string]$ImageFormat = 'png',

    [switch]$Force,
    [switch]$ShowWord,
    [switch]$KeepOpen,

    [ValidateRange(1, 5)]
    [int]$Attempts = 2,

    [ValidateRange(0, 86400)]
    [int]$TimeoutSeconds = 120,

    [switch]$SkipPdfWarningFix,
    [switch]$Unblock
)
Set-StrictMode -Version Latest

# The implementation lives in PdfToWord.Core.ps1, shared with PdfToWordWorker.ps1.
. (Join-Path $PSScriptRoot 'PdfToWord.Core.ps1')

try {
    Invoke-PdfToWordConversion -Path $Path -OutPath $OutPath -Format $Format `
        -ImagesFolderName $ImagesFolderName -ImageFormat $ImageFormat `
        -Force:$Force -ShowWord:$ShowWord -KeepOpen:$KeepOpen `
        -Attempts $Attempts -TimeoutSeconds $TimeoutSeconds `
        -SkipPdfWarningFix:$SkipPdfWarningFix -Unblock:$Unblock
}
catch {
    if ($_.Exception -is [System.ArgumentException]) { throw $_.Exception.Message }
    throw
}
