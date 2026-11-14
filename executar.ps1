param(
    [string]$Python = "py",
    [string]$Period = "2026-10",
    [switch]$NoCache,
    [switch]$NoOCR,
    [switch]$Strict,
    [switch]$Validate,
    [ValidateSet("off", "consultivo", "automatico")][string]$AiMode,
    [ValidateSet("excecoes", "todos")][string]$AiScope
)
$ErrorActionPreference = "Stop"
Push-Location $PSScriptRoot
try {
    if (-not $env:GEMINI_API_KEY -and -not (Test-Path -LiteralPath ".env")) {
        $env:GEMINI_API_KEY = [Environment]::GetEnvironmentVariable('GEMINI_API_KEY', 'User')
    }
    [string[]]$pythonArgs = if ($Python -eq "py") { @("-3.12") } else { @() }
    & $Python @pythonArgs -c "import sys; assert (3,11) <= sys.version_info[:2] < (3,14), 'Use Python 3.11, 3.12 ou 3.13'"
    if ($LASTEXITCODE -ne 0) { throw "Python compativel nao encontrado. Instale Python 3.12 ou informe -Python com o caminho do executavel." }
    & $Python @pythonArgs -m pip install -r requirements-lock.txt
    if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar dependencias no Python selecionado." }
    if ($Validate) {
        & $Python @pythonArgs -m pytest -q
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    }
    $pipelineArgs = @("-m", "goener", "--config", (Join-Path $PSScriptRoot "config.toml"), "--period", $Period)
    if ($NoCache) { $pipelineArgs += "--no-cache" }
    if ($NoOCR) { $pipelineArgs += "--no-ocr" }
    if ($Strict) { $pipelineArgs += "--strict" }
    if ($AiMode) { $pipelineArgs += @("--ai-mode", $AiMode) }
    if ($AiScope) { $pipelineArgs += @("--ai-scope", $AiScope) }
    & $Python @pythonArgs @pipelineArgs
    $pipelineCode = $LASTEXITCODE
} finally {
    Pop-Location
}
exit $pipelineCode
