param(
    [string]$Python = "python",
    [string]$Version = "v1.0.0"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvDir = Join-Path $projectRoot ".venv-build"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$source = Join-Path $projectRoot "採購單大量修改數量與總價.py"
$distDir = Join-Path $projectRoot "dist"
$buildDir = Join-Path $projectRoot "build"
$appName = "採購單數量與金額調整工具"
$assetBaseName = "purchase-order-quantity-updater"

if (-not (Test-Path -LiteralPath $venvPython)) {
    & $Python -m venv $venvDir
    if ($LASTEXITCODE -ne 0) { throw "建立建置環境失敗" }
}

& $venvPython -m pip install --disable-pip-version-check -r (Join-Path $projectRoot "requirements-build.txt")
if ($LASTEXITCODE -ne 0) { throw "安裝建置套件失敗" }

& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --collect-all openpyxl `
    --name $appName `
    --distpath $distDir `
    --workpath $buildDir `
    --specpath $buildDir `
    $source
if ($LASTEXITCODE -ne 0) { throw "建立執行檔失敗" }

$releaseDir = Join-Path $distDir "$assetBaseName-$Version-win64"
if (Test-Path -LiteralPath $releaseDir) {
    $resolvedDist = [System.IO.Path]::GetFullPath($distDir).TrimEnd('\') + '\'
    $resolvedRelease = [System.IO.Path]::GetFullPath($releaseDir)
    if (-not $resolvedRelease.StartsWith($resolvedDist, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "拒絕清理預期範圍外的路徑：$resolvedRelease"
    }
    [System.IO.Directory]::Delete($resolvedRelease, $true)
}

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $distDir "$appName.exe") -Destination $releaseDir
Copy-Item -LiteralPath (Join-Path $projectRoot "出貨清單_範本.xlsx") -Destination $releaseDir
Copy-Item -LiteralPath (Join-Path $projectRoot "使用說明.txt") -Destination $releaseDir

$exeAssetPath = Join-Path $distDir "$assetBaseName.exe"
Copy-Item -LiteralPath (Join-Path $distDir "$appName.exe") -Destination $exeAssetPath -Force

$zipPath = Join-Path $distDir "$assetBaseName-$Version-win64.zip"
if (Test-Path -LiteralPath $zipPath) {
    [System.IO.File]::Delete([System.IO.Path]::GetFullPath($zipPath))
}
Compress-Archive -Path (Join-Path $releaseDir "*") -DestinationPath $zipPath -CompressionLevel Optimal

Get-FileHash -Algorithm SHA256 -LiteralPath $exeAssetPath, $zipPath |
    Select-Object Path, Hash
