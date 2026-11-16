param(
    [Parameter(Mandatory = $true)]
    [string[]]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,
    [ValidateRange(96, 400)]
    [int]$ResolutionDpi = 200,
    [ValidateRange(1, 100000000)]
    [long]$MaxPixelsPerPage = 40000000
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop

function Await-WinRTOperation {
    param([Parameter(Mandatory = $true)]$Operation, [Parameter(Mandatory = $true)][Type]$ResultType)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
        } | Select-Object -First 1
    if (-not $method) { throw 'Windows Runtime async-operation bridge is unavailable.' }
    $task = $method.MakeGenericMethod(@($ResultType)).Invoke($null, @($Operation))
    return $task.GetAwaiter().GetResult()
}

function Await-WinRTAction {
    param([Parameter(Mandatory = $true)]$Operation)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
        } | Select-Object -First 1
    if (-not $method) { throw 'Windows Runtime async-action bridge is unavailable.' }
    $task = $method.Invoke($null, @($Operation))
    $task.GetAwaiter().GetResult()
}

if ($PSVersionTable.PSVersion.Major -ne 5) {
    throw 'This helper requires Windows PowerShell 5.1.'
}
if (Test-Path -LiteralPath $OutputDirectory) {
    throw 'Output directory already exists; choose a new directory to avoid overwriting files.'
}

$resolvedInputs = @()
foreach ($path in $InputPath) {
    $resolved = (Resolve-Path -LiteralPath $path -ErrorAction Stop).Path
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { throw "Input is not a file: $resolved" }
    if ([System.IO.Path]::GetExtension($resolved).ToLowerInvariant() -ne '.pdf') {
        throw "Unsupported PDF extension: $resolved"
    }
    $resolvedInputs += $resolved
}
if ($resolvedInputs.Count -eq 0) { throw 'Provide at least one PDF file.' }
$duplicateNames = $resolvedInputs |
    ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_).ToLowerInvariant() } |
    Group-Object | Where-Object { $_.Count -gt 1 }
if ($duplicateNames) { throw 'Input PDF base names must be unique because output text files use those names.' }

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType=WindowsRuntime]

$language = [Windows.Globalization.Language]::new('zh-Hans-CN')
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
if (-not $engine) { throw 'The Windows OCR engine for zh-Hans-CN is unavailable.' }

$utf8 = New-Object System.Text.UTF8Encoding($true)
New-Item -ItemType Directory -Path $OutputDirectory -ErrorAction Stop | Out-Null
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path
$records = @()
$totalPages = 0

foreach ($source in $resolvedInputs) {
    $sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($source)
    $textPath = Join-Path $resolvedOutput ($baseName + '.txt')
    if (Test-Path -LiteralPath $textPath) { throw "Output file already exists: $textPath" }

    $file = Await-WinRTOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($source)) ([Windows.Storage.StorageFile])
    $document = $null
    $pageManifest = @()
    $textBuilder = New-Object System.Text.StringBuilder
    try {
        $document = Await-WinRTOperation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
        if ($document.PageCount -lt 1) { throw "PDF has no pages: $source" }
        for ($index = 0; $index -lt $document.PageCount; $index++) {
            $page = $null
            $stream = $null
            $bitmap = $null
            try {
                $page = $document.GetPage([uint32]$index)
                $scale = [double]$ResolutionDpi / 96.0
                $width = [uint32][Math]::Ceiling($page.Size.Width * $scale)
                $height = [uint32][Math]::Ceiling($page.Size.Height * $scale)
                $pixelCount = [long]$width * [long]$height
                if ($width -lt 1 -or $height -lt 1 -or $pixelCount -gt $MaxPixelsPerPage) {
                    throw "Rendered page $($index + 1) exceeds the $MaxPixelsPerPage-pixel limit; lower ResolutionDpi or review this PDF separately."
                }
                $options = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
                $options.DestinationWidth = $width
                $options.DestinationHeight = $height
                $stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
                $null = Await-WinRTAction ($page.RenderToStreamAsync($stream, $options))
                $null = $stream.Seek(0)
                $decoder = Await-WinRTOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
                $bitmap = Await-WinRTOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
                $result = Await-WinRTOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
                $pageLines = @($result.Lines | ForEach-Object { $_.Text })
                [void]$textBuilder.AppendLine(('--- Page {0} ---' -f ($index + 1)))
                foreach ($line in $pageLines) { [void]$textBuilder.AppendLine([string]$line) }
                $pageManifest += [PSCustomObject]@{
                    page_number = $index + 1
                    width = $bitmap.PixelWidth
                    height = $bitmap.PixelHeight
                    line_count = $pageLines.Count
                }
            }
            finally {
                if ($bitmap) { $bitmap.Dispose() }
                if ($stream) { $stream.Dispose() }
                if ($page) { $page.Dispose() }
            }
        }
    }
    finally {
        # PdfDocument exposes no Dispose/Close method in the Windows PowerShell 5.1 WinRT projection.
    }

    $currentHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($currentHash -ne $sourceHash) { throw "Source PDF changed while processing; no text written for $source" }
    [System.IO.File]::WriteAllText($textPath, $textBuilder.ToString(), $utf8)
    $totalPages += $pageManifest.Count
    $records += [PSCustomObject]@{
        source_path = $source
        source_sha256 = $sourceHash
        output_text = [System.IO.Path]::GetFileName($textPath)
        language = 'zh-Hans-CN'
        resolution_dpi = $ResolutionDpi
        page_count = $pageManifest.Count
        pages = $pageManifest
        status = 'transcribed_requires_human_review'
    }
}

$manifestPath = Join-Path $resolvedOutput 'ocr_manifest.json'
if (Test-Path -LiteralPath $manifestPath) { throw "Manifest already exists: $manifestPath" }
$records | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Output ("Transcribed {0} PDF file(s), {1} page(s). Review every output against its source; this helper does not verify OCR accuracy." -f $records.Count, $totalPages)
Write-Output $resolvedOutput




