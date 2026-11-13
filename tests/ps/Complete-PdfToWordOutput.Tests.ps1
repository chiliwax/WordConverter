<#
    Pester tests for PdfToWord.Core.ps1. No Microsoft Word is required.

    Run with:
        Invoke-Pester -Script .\tests\ps\Complete-PdfToWordOutput.Tests.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $repoRoot 'PdfToWord.Core.ps1')

Add-Type -AssemblyName System.Drawing

function New-TestDir {
    $path = Join-Path ([IO.Path]::GetTempPath()) ("pdf2word-pester-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    return $path
}

function Remove-TestDir {
    param([string]$Path)
    Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
}

function New-TestImage {
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Png', 'Jpeg', 'Gif', 'Bmp')][string]$Format = 'Png'
    )
    $bitmap = New-Object System.Drawing.Bitmap -ArgumentList 4, 4
    try {
        $imageFormat = switch ($Format) {
            'Png'  { [System.Drawing.Imaging.ImageFormat]::Png }
            'Jpeg' { [System.Drawing.Imaging.ImageFormat]::Jpeg }
            'Gif'  { [System.Drawing.Imaging.ImageFormat]::Gif }
            'Bmp'  { [System.Drawing.Imaging.ImageFormat]::Bmp }
        }
        $bitmap.Save($Path, $imageFormat)
    }
    finally {
        $bitmap.Dispose()
    }
}

function New-HtmlWithImages {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$Stem,
        [Parameter(Mandatory)][string]$FolderName,
        [string[]]$Images = @('image001.png')
    )
    $htmlPath = Join-Path $Directory ($Stem + '.html')
    $imageDir = Join-Path $Directory $FolderName
    New-Item -ItemType Directory -Path $imageDir -Force | Out-Null
    $refs = @($Images | ForEach-Object { "$FolderName/$_" })
    $html = '<html><body>' + (($refs | ForEach-Object { "<img src=""$_"">" }) -join '') + '</body></html>'
    [IO.File]::WriteAllText($htmlPath, $html, (New-Object Text.UTF8Encoding($false)))
    return @{ Html = $htmlPath; Images = $imageDir }
}

Describe 'Complete-PdfToWordOutput HTML companion folders' {
    It 'renames an English _files folder to source_folder' {
        $dir = New-TestDir
        try {
            $test = New-HtmlWithImages -Directory $dir -Stem 'doc' -FolderName 'doc_files'
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $dir 'source_folder') -PathType Container) | Should Be $true
            (Test-Path $test.Images) | Should Be $false
            (Test-Path (Join-Path $dir 'source_folder\image001.png')) | Should Be $true
            ([IO.File]::ReadAllText($test.Html)) | Should Match 'source_folder/image001\.png'
            ([IO.File]::ReadAllText($test.Html)) | Should Not Match 'doc_files'
        }
        finally { Remove-TestDir $dir }
    }

    It 'renames a French _fichiers folder with an encoded space' {
        $dir = New-TestDir
        try {
            $test = New-HtmlWithImages -Directory $dir -Stem 'NP 2025-32' -FolderName 'NP 2025-32_fichiers'
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png
            $html = [IO.File]::ReadAllText($test.Html).Replace('NP 2025-32_fichiers', 'NP%202025-32_fichiers')
            [IO.File]::WriteAllText($test.Html, $html, (New-Object Text.UTF8Encoding($false)))

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $dir 'source_folder') -PathType Container) | Should Be $true
            (Test-Path $test.Images) | Should Be $false
            ([IO.File]::ReadAllText($test.Html)) | Should Match 'source_folder/image001\.png'
            ([IO.File]::ReadAllText($test.Html)) | Should Not Match 'NP%202025-32_fichiers'
        }
        finally { Remove-TestDir $dir }
    }

    It 'renames a folder whose link mixes an encoded space and a raw accented letter' {
        $dir = New-TestDir
        try {
            $accented = [char]0xE9
            $stem = "Mesure applicative_FR_sign${accented}"
            $folder = "${stem}_fichiers"
            $test = New-HtmlWithImages -Directory $dir -Stem $stem -FolderName $folder
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png

            # Exactly what Word wrote for the reported document: the space is
            # percent-encoded, the accented letter is left raw.
            $encoded = 'Mesure%20applicative_FR_sign' + $accented + '_fichiers'
            $html = [IO.File]::ReadAllText($test.Html).Replace($folder, $encoded)
            [IO.File]::WriteAllText($test.Html, $html, (New-Object Text.UTF8Encoding($false)))

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $dir 'source_folder\image001.png')) | Should Be $true
            $patched = [IO.File]::ReadAllText($test.Html)
            $patched | Should Match 'source_folder/image001\.png'
            $patched | Should Not Match ([regex]::Escape($encoded))
        }
        finally { Remove-TestDir $dir }
    }

    It 'renames a folder whose name is HTML-escaped in the link' {
        $dir = New-TestDir
        try {
            $test = New-HtmlWithImages -Directory $dir -Stem 'A & B' -FolderName 'A & B_files'
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png

            # "A & B_files" as Word may serialize it in an attribute: %20 for the
            # spaces, &amp; for the ampersand.
            $encoded = 'A%20&amp;%20B_files'
            $html = [IO.File]::ReadAllText($test.Html).Replace('A & B_files', $encoded)
            [IO.File]::WriteAllText($test.Html, $html, (New-Object Text.UTF8Encoding($false)))

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $dir 'source_folder\image001.png')) | Should Be $true
            $patched = [IO.File]::ReadAllText($test.Html)
            $patched | Should Match 'source_folder/image001\.png'
            $patched | Should Not Match ([regex]::Escape($encoded))
        }
        finally { Remove-TestDir $dir }
    }

    It 'renames a folder whose name contains parentheses' {
        $dir = New-TestDir
        try {
            $stem = 'ARCHER DESIGN PATTERNS AWS S3 (EN)'
            $folder = $stem + '_fichiers'
            $test = New-HtmlWithImages -Directory $dir -Stem $stem -FolderName $folder
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png

            # Word leaves the parentheses raw and encodes only the spaces.
            $encoded = 'ARCHER%20DESIGN%20PATTERNS%20AWS%20S3%20(EN)_fichiers'
            $html = [IO.File]::ReadAllText($test.Html).Replace($folder, $encoded)
            [IO.File]::WriteAllText($test.Html, $html, (New-Object Text.UTF8Encoding($false)))

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $dir 'source_folder\image001.png')) | Should Be $true
            $patched = [IO.File]::ReadAllText($test.Html)
            $patched | Should Match 'source_folder/image001\.png'
            $patched | Should Not Match ([regex]::Escape($encoded))
        }
        finally { Remove-TestDir $dir }
    }

    It 'renames a folder whose name contains an apostrophe' {
        $dir = New-TestDir
        try {
            $stem = "CAM0036-V2-Origine de l'Etablissement"
            $folder = $stem + '_fichiers'
            $test = New-HtmlWithImages -Directory $dir -Stem $stem -FolderName $folder
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png

            $encoded = "CAM0036-V2-Origine%20de%20l'Etablissement_fichiers"
            $html = [IO.File]::ReadAllText($test.Html).Replace($folder, $encoded)
            [IO.File]::WriteAllText($test.Html, $html, (New-Object Text.UTF8Encoding($false)))

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $dir 'source_folder\image001.png')) | Should Be $true
            $patched = [IO.File]::ReadAllText($test.Html)
            $patched | Should Match 'source_folder/image001\.png'
            $patched | Should Not Match ([regex]::Escape($encoded))
        }
        finally { Remove-TestDir $dir }
    }

    It 'does not touch a stale folder older than -Since' {
        $dir = New-TestDir
        try {
            $test = New-HtmlWithImages -Directory $dir -Stem 'doc' -FolderName 'doc_fichiers'
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png
            (Get-Item -LiteralPath $test.Images).LastWriteTime = (Get-Date).AddHours(-2)

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png `
                -Since (Get-Date) -Force

            (Test-Path $test.Images) | Should Be $true
            (Test-Path (Join-Path $dir 'source_folder')) | Should Be $false
        }
        finally { Remove-TestDir $dir }
    }

    It 'replaces an existing source_folder with -Force' {
        $dir = New-TestDir
        try {
            $test = New-HtmlWithImages -Directory $dir -Stem 'doc' -FolderName 'doc_fichiers'
            New-TestImage -Path (Join-Path $test.Images 'image001.png') -Format Png
            $existing = Join-Path $dir 'source_folder'
            New-Item -ItemType Directory -Path $existing -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $existing 'old.txt') -Value 'old'

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            (Test-Path (Join-Path $existing 'old.txt')) | Should Be $false
            (Test-Path (Join-Path $existing 'image001.png')) | Should Be $true
            (Test-Path $test.Images) | Should Be $false
        }
        finally { Remove-TestDir $dir }
    }

    It 'converts jpg, gif and bmp files to png and repairs links' {
        $dir = New-TestDir
        try {
            $test = New-HtmlWithImages -Directory $dir -Stem 'doc' -FolderName 'doc_files' `
                -Images @('image001.jpg', 'image002.gif', 'image003.bmp')
            New-TestImage -Path (Join-Path $test.Images 'image001.jpg') -Format Jpeg
            New-TestImage -Path (Join-Path $test.Images 'image002.gif') -Format Gif
            New-TestImage -Path (Join-Path $test.Images 'image003.bmp') -Format Bmp

            Complete-PdfToWordOutput -Spec $script:PdfToWordFormats['html'] `
                -OutputPath $test.Html -ImagesFolderName 'source_folder' -ImageFormat png -Force

            $folder = Join-Path $dir 'source_folder'
            foreach ($name in 'image001.png', 'image002.png', 'image003.png') {
                (Test-Path (Join-Path $folder $name)) | Should Be $true
            }
            foreach ($name in 'image001.jpg', 'image002.gif', 'image003.bmp') {
                (Test-Path (Join-Path $folder $name)) | Should Be $false
            }
            $html = [IO.File]::ReadAllText($test.Html)
            $html | Should Match 'source_folder/image001\.png'
            $html | Should Match 'source_folder/image002\.png'
            $html | Should Match 'source_folder/image003\.png'
        }
        finally { Remove-TestDir $dir }
    }
}

Describe 'Resolve-PdfToWordJob validation' {
    It 'rejects a folder-shaped output path' {
        $dir = New-TestDir
        try {
            $pdf = Join-Path $dir 'in.pdf'
            [IO.File]::WriteAllText($pdf, '%PDF-1.4')
            $threw = $false
            try {
                Resolve-PdfToWordJob -Path $pdf -OutPath ($dir + '\') -Format docx
            }
            catch { $threw = $true }
            $threw | Should Be $true
        }
        finally { Remove-TestDir $dir }
    }

    It 'refuses an existing output without -Force' {
        $dir = New-TestDir
        try {
            $pdf = Join-Path $dir 'in.pdf'
            $out = Join-Path $dir 'out.docx'
            [IO.File]::WriteAllText($pdf, '%PDF-1.4')
            [IO.File]::WriteAllText($out, 'existing')
            $threw = $false
            try {
                Resolve-PdfToWordJob -Path $pdf -OutPath $out -Format docx
            }
            catch { $threw = $true }
            $threw | Should Be $true
        }
        finally { Remove-TestDir $dir }
    }

    It 'accepts an existing output with -Force' {
        $dir = New-TestDir
        try {
            $pdf = Join-Path $dir 'in.pdf'
            $out = Join-Path $dir 'out.docx'
            [IO.File]::WriteAllText($pdf, '%PDF-1.4')
            [IO.File]::WriteAllText($out, 'existing')
            $job = Resolve-PdfToWordJob -Path $pdf -OutPath $out -Format docx -Force
            $job.OutFile | Should Be $out
            $job.Format | Should Be 'docx'
        }
        finally { Remove-TestDir $dir }
    }
}
