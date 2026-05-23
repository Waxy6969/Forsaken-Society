$ErrorActionPreference = "Stop"

$python = "C:\Users\ADMIN\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
  $python = "python"
}

& $python "$PSScriptRoot\app.py"
