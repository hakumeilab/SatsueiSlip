$ErrorActionPreference = "Stop"

$rootDir = Split-Path -Parent $PSScriptRoot
$specPath = Join-Path $rootDir "SatsueiSlip.spec"
$issPath = Join-Path $rootDir "installer\SatsueiSlip.iss"

python -m pip install pyinstaller
pyinstaller $specPath --noconfirm

$isccPath = $null
$isccCommand = Get-Command iscc -ErrorAction SilentlyContinue
if ($isccCommand) {
    $isccPath = $isccCommand.Source
}

if (-not $isccPath) {
    $defaultPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (Test-Path -LiteralPath $defaultPath) {
        $isccPath = $defaultPath
    }
}

if (-not $isccPath) {
    throw "ISCC.exe が見つかりません。Inno Setup 6 をインストールし、iscc に PATH を通してください。"
}

& $isccPath $issPath
