param([Parameter(Mandatory=$true)][string]$ImagePath)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime] | Out-Null

$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1

function Wait-Operation($operation, [type]$resultType) {
    $task = $asTask.MakeGenericMethod($resultType).Invoke($null, @($operation))
    $task.GetAwaiter().GetResult()
}

$language = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages |
    Where-Object LanguageTag -eq 'zh-Hans-CN' | Select-Object -First 1
if (!$language) { throw 'Windows OCR language zh-Hans-CN is missing.' }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
$memory = [IO.MemoryStream]::new([IO.File]::ReadAllBytes($ImagePath))
$stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($memory)
try {
    $decoder = Wait-Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Wait-Operation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $result = Wait-Operation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines = @(foreach ($line in $result.Lines) {
            $words = @($line.Words)
            $rect = $words[0].BoundingRect
            @{ text = ($line.Words.Text -join ''); x = $rect.X; y = $rect.Y; h = $rect.Height }
        })
        ConvertTo-Json -InputObject $lines -Compress
    } finally { $bitmap.Dispose() }
} finally {
    $stream.Dispose()
    $memory.Dispose()
}
