[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$CommandArgs
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PaddlePython = Join-Path $ProjectDir ".paddle-venv\Scripts\python.exe"
$PaddleHome = Join-Path $ProjectDir ".paddle-home"
$PaddleCache = Join-Path $ProjectDir ".paddle-cache"
$PaddleTemp = Join-Path $ProjectDir ".paddle-tmp"

if (-not (Test-Path $PaddlePython)) {
    throw "PaddleOCR environment not found. Run .\setup_paddleocr.ps1 first."
}

foreach ($directory in @(
    $PaddleHome,
    $PaddleCache,
    (Join-Path $PaddleCache "paddlex"),
    (Join-Path $PaddleCache "modelscope"),
    (Join-Path $PaddleCache "huggingface"),
    $PaddleTemp
)) {
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
}

$env:HOME = $PaddleHome
$env:USERPROFILE = $PaddleHome
$env:PADDLE_PDX_CACHE_HOME = Join-Path $PaddleCache "paddlex"
$env:MODELSCOPE_CACHE = Join-Path $PaddleCache "modelscope"
$env:HF_HOME = Join-Path $PaddleCache "huggingface"
$env:HUGGINGFACE_HUB_CACHE = Join-Path $env:HF_HOME "hub"
$env:HF_HUB_CACHE = $env:HUGGINGFACE_HUB_CACHE
$env:TEMP = $PaddleTemp
$env:TMP = $PaddleTemp
# The first-run connectivity probe can hang behind the corporate network;
# model downloads still require network access when a model is not cached.
$env:PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = "True"

Push-Location $ProjectDir
try {
    if (-not $CommandArgs -or $CommandArgs.Count -eq 0) {
        & $PaddlePython -u -c "import os, paddle, paddleocr; print('paddle=' + paddle.__version__); print('paddleocr=' + getattr(paddleocr, '__version__', 'unknown')); print('cache=' + os.environ['PADDLE_PDX_CACHE_HOME'])"
    } else {
        & $PaddlePython @CommandArgs
    }
    $exitCode = $LASTEXITCODE
} finally {
    Pop-Location
}
exit $exitCode
