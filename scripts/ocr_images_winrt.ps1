param(
    [Parameter(Mandatory = $true)]
    [string[]]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'

function Await-WinRTOperation {
    param(
        [Parameter(Mandatory = $true)]$Operation,
        [Parameter(Mandatory = $true)][Type]$ResultType
    )

    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq 'AsTask' -and
            $_.IsGenericMethodDefinition -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
        } |
        Select-Object -First 1

    if (-not $method) {
        throw 'Windows Runtime async bridge is unavailable.'
    }

    $task = $method.MakeGenericMethod(@($ResultType)).Invoke($null, @($Operation))
    return $task.GetAwaiter().GetResult()
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
    if ([System.IO.Path]::GetExtension($resolved).ToLowerInvariant() -notin @('.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff')) {
        throw "Unsupported image extension: $resolved"
    }
    $resolvedInputs += $resolved
}

if ($resolvedInputs.Count -eq 0) {
    throw 'Provide at least one image file.'
}

$duplicateNames = $resolvedInputs |
    ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_).ToLowerInvariant() } |
    Group-Object |
    Where-Object { $_.Count -gt 1 }
if ($duplicateNames) {
    throw 'Input image base names must be unique because output text files use those names.'
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime]

$language = [Windows.Globalization.Language]::new('zh-Hans-CN')
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
if (-not $engine) {
    throw 'The Windows OCR engine for zh-Hans-CN is unavailable.'
}

New-Item -ItemType Directory -Path $OutputDirectory -ErrorAction Stop | Out-Null
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path
$records = @()
$utf8 = New-Object System.Text.UTF8Encoding($true)

foreach ($source in $resolvedInputs) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($source)
    $textPath = Join-Path $resolvedOutput ($baseName + '.txt')
    if (Test-Path -LiteralPath $textPath) {
        throw "Output file already exists: $textPath"
    }

    $file = Await-WinRTOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($source)) ([Windows.Storage.StorageFile])
    $stream = Await-WinRTOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Await-WinRTOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Await-WinRTOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $result = Await-WinRTOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lineTexts = @($result.Lines | ForEach-Object { $_.Text })
        $recognizedText = [string]::Join([Environment]::NewLine, $lineTexts)
        [System.IO.File]::WriteAllText($textPath, $recognizedText, $utf8)

        $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
        $records += [PSCustomObject]@{
            source_path = $source
            source_sha256 = $hash
            output_text = [System.IO.Path]::GetFileName($textPath)
            language = 'zh-Hans-CN'
            width = $bitmap.PixelWidth
            height = $bitmap.PixelHeight
            line_count = $lineTexts.Count
            status = 'transcribed_requires_human_review'
        }
    }
    finally {
        if ($stream) { $stream.Dispose() }
    }
}

$manifestPath = Join-Path $resolvedOutput 'ocr_manifest.json'
if (Test-Path -LiteralPath $manifestPath) {
    throw "Manifest already exists: $manifestPath"
}
$records | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Output ("Transcribed {0} image(s). Review every output against its source; this helper does not verify OCR accuracy." -f $records.Count)
Write-Output $resolvedOutput
