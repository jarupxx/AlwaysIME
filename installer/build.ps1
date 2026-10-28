# AlwaysIME installer build script (run on Windows)
# Requires: .NET 10 SDK, Inno Setup 6.3 or later, including 7.x (ISCC.exe)
# Usage:    pwsh .\installer\build.ps1            (or: powershell -File .\installer\build.ps1)
#           .\installer\build.ps1 -Iscc "D:\Tools\Inno Setup 6\ISCC.exe"

param(
    [string]$Iscc = ""
)

$ErrorActionPreference = "Stop"
$root    = Split-Path -Parent $PSScriptRoot
$csproj  = Join-Path $root "AlwaysIME\AlwaysIME.csproj"
$publish = Join-Path $PSScriptRoot "publish"
$script  = Join-Path $PSScriptRoot "AlwaysIME.iss"

# Version comes from VersionPrefix in the csproj (single source of truth)
[xml]$xml = Get-Content -LiteralPath $csproj -Raw
$version = ($xml.Project.PropertyGroup | ForEach-Object { $_.VersionPrefix } | Where-Object { $_ } | Select-Object -First 1)
if (-not $version) { throw "VersionPrefix was not found in $csproj" }
Write-Host "Version: $version"

# Locate ISCC.exe
if (-not $Iscc) {
    # Newer major version first (7 before 6); then whatever is on PATH
    $candidates = @()
    foreach ($ver in @("7", "6")) {
        $candidates += (Join-Path $env:ProgramFiles "Inno Setup $ver\ISCC.exe")
        $candidates += (Join-Path ${env:ProgramFiles(x86)} "Inno Setup $ver\ISCC.exe")
        $candidates += (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup $ver\ISCC.exe")
    }
    $Iscc = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $Iscc) {
        $cmd = Get-Command ISCC.exe -ErrorAction SilentlyContinue
        if ($cmd) { $Iscc = $cmd.Source }
    }
}
if (-not $Iscc -or -not (Test-Path $Iscc)) { throw "ISCC.exe was not found. Install Inno Setup or pass -Iscc." }

# Publish (framework-dependent, x64). The .NET 10 Desktop Runtime is required on the target PC.
if (Test-Path $publish) { Remove-Item -LiteralPath $publish -Recurse -Force }
dotnet publish $csproj -c Release -r win-x64 --self-contained false -o $publish
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed" }

# Required files must be present before packaging
foreach ($f in @("AlwaysIME.exe", "AlwaysIME.dll", "AlwaysIME.runtimeconfig.json", "AlwaysIME.dll.config.default", "LICENSE.txt")) {
    if (-not (Test-Path (Join-Path $publish $f))) { throw "Missing in publish output: $f" }
}

# Build the installer
Write-Host "ISCC: $Iscc"
& $Iscc "/DMyAppVersion=$version" "/DPublishDir=publish" $script
if ($LASTEXITCODE -ne 0) { throw "ISCC failed" }

$setup = Join-Path $PSScriptRoot "Output\AlwaysIME-$version-x64-setup.exe"
Write-Host ""
Write-Host "Installer: $setup"
Write-Host "SHA256:    $((Get-FileHash -LiteralPath $setup -Algorithm SHA256).Hash)"
