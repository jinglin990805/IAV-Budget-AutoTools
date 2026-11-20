[CmdletBinding()]
param(
    [string]$PypiIndex = "https://pypi.tuna.tsinghua.edu.cn/simple"
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PaddleEnvDir = Join-Path $ProjectDir ".paddle-venv"
$PaddlePython = Join-Path $PaddleEnvDir "Scripts\python.exe"
$PaddleHome = Join-Path $ProjectDir ".paddle-home"
$PaddleCache = Join-Path $ProjectDir ".paddle-cache"
$PaddleTemp = Join-Path $ProjectDir ".paddle-tmp"

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

if (-not (Test-Path $PaddlePython)) {
    if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
        throw "Python launcher 'py' was not found. Install Python 3.12 first."
    }
    & py -3.12 -m venv --without-pip $PaddleEnvDir
}

if (-not (Test-Path $PaddlePython)) {
    throw "Failed to create the Python 3.12 environment at $PaddleEnvDir"
}

# pip --python works with the Windows Store Python even when ensurepip cannot
# write its temporary wheel in the managed profile.
$PipTargetArgs = @("-3.12", "-m", "pip", "--python", $PaddlePython)
function Invoke-PaddlePip {
    param([string[]]$Arguments)
    & py @PipTargetArgs @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "pip command failed with exit code ${LASTEXITCODE}: $($Arguments -join ' ')"
    }
}

Invoke-PaddlePip @("install", "--upgrade", "pip", "--disable-pip-version-check")
Invoke-PaddlePip @(
    "install", "--no-cache-dir", "paddlepaddle==3.2.0",
    "-i", "https://www.paddlepaddle.org.cn/packages/stable/cpu/",
    "--disable-pip-version-check"
)
Invoke-PaddlePip @(
    "install", "--no-cache-dir", "paddleocr==3.7.0",
    "-i", $PypiIndex, "--disable-pip-version-check"
)

# PaddlePaddle/PaddleX/model downloaders otherwise default to the managed user
# profile, which is not writable in this workstation environment. Set these
# only after the `py` launcher has finished managing the target environment.
$env:HOME = $PaddleHome
$env:USERPROFILE = $PaddleHome
$env:PADDLE_PDX_CACHE_HOME = Join-Path $PaddleCache "paddlex"
$env:MODELSCOPE_CACHE = Join-Path $PaddleCache "modelscope"
$env:HF_HOME = Join-Path $PaddleCache "huggingface"
$env:HUGGINGFACE_HUB_CACHE = Join-Path $env:HF_HOME "hub"
$env:HF_HUB_CACHE = $env:HUGGINGFACE_HUB_CACHE
$env:TEMP = $PaddleTemp
$env:TMP = $PaddleTemp
$env:PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = "True"

& $PaddlePython -u -c "import paddle, paddleocr, cv2; print('paddle=' + paddle.__version__); print('paddleocr=' + getattr(paddleocr, '__version__', 'unknown')); print('cv2=' + cv2.__version__)"
if ($LASTEXITCODE -ne 0) {
    throw "PaddleOCR import verification failed"
}

Write-Host "PaddleOCR setup completed in $PaddleEnvDir"
Write-Host "Use .\paddleocr_env.ps1 ocr_benchmark.py .playwright-mcp\ocr-fixtures --limit 4 --dpi 150 --max-pages 1"
