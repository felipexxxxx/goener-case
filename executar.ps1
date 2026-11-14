param(
    [string]$Python = "python",
    [string]$Period = "2026-10",
    [switch]$NoCache,
    [switch]$NoOCR,
    [switch]$Strict,
    [switch]$Validate
)
$ErrorActionPreference = "Stop"
Push-Location $PSScriptRoot
try {
    $venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $venvPython)) {
        & $Python -c "import sys; assert (3,11) <= sys.version_info[:2] < (3,14), 'Use Python 3.11, 3.12 ou 3.13'"
        if ($LASTEXITCODE -ne 0) { throw "Python compativel nao encontrado. Instale Python 3.12 ou informe -Python com o caminho do executavel." }
        & $Python -m venv .venv
        if ($LASTEXITCODE -ne 0) { throw "Falha ao criar o ambiente virtual." }
    }
    $lockHash = (Get-FileHash -LiteralPath requirements-lock.txt -Algorithm SHA256).Hash
    $lockHash += (Get-FileHash -LiteralPath pyproject.toml -Algorithm SHA256).Hash
    $setupMarker = Join-Path $PSScriptRoot ".venv\goener-setup.sha256"
    $installedHash = if (Test-Path -LiteralPath $setupMarker) { (Get-Content -LiteralPath $setupMarker -Raw).Trim() } else { "" }
    if ($installedHash -ne $lockHash) {
        & $venvPython -m pip install -r requirements-lock.txt
        if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar dependencias." }
        & $venvPython -m pip install -e . --no-deps --no-build-isolation
        if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar o projeto." }
        Set-Content -LiteralPath $setupMarker -Value $lockHash -Encoding ascii
    }
    if ($Validate) {
        & $venvPython -m pytest -q
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    }
    $pipelineArgs = @("-m", "goener", "--config", (Join-Path $PSScriptRoot "config.toml"), "--period", $Period)
    if ($NoCache) { $pipelineArgs += "--no-cache" }
    if ($NoOCR) { $pipelineArgs += "--no-ocr" }
    if ($Strict) { $pipelineArgs += "--strict" }
    & $venvPython @pipelineArgs
    $pipelineCode = $LASTEXITCODE
} finally {
    Pop-Location
}
exit $pipelineCode
