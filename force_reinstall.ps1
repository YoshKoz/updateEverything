# Force remove broken MSI registrations and reinstall via winget
# Run as Administrator

$guids = @{
    'GitHub CLI 2.88.1'  = '{399408A8-1AA3-45A2-B4C8-905FAD1862F9}'
    'Neovim 0.11.6'      = '{716666C3-6EF3-4BE7-BD7D-B872DF79F785}'
    'CMake 4.2.3'        = '{7C40E3FE-8918-49C8-95D4-F6BF883C74DF}'
    'calibre 9.5.0'      = '{92FBC5AA-50A3-48E7-A458-2B78563B3993}'
    'calibre 9.6.0'      = '{C914F9C8-A68C-460F-8DCE-353983D5B7FB}'
}

# Remove registry uninstall entries so new installer doesn't find old version
foreach ($name in $guids.Keys) {
    $guid = $guids[$name]
    $guidClean = $guid.Trim('{}')
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$guid",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$guid",
        "HKLM:\SOFTWARE\Classes\Installer\Products\$($guidClean -replace '(.)(.)(.)(.)(.)(.)(.)(.)-(..)(..)-(..)(..)-(..)(..)(..)(..)(..)(..)(..)(..)', '$8$7$6$5$4$3$2$1$10$9$12$11$14$13$24$23$22$21$20$19$18$17$16$15')"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            Remove-Item $p -Recurse -Force
            Write-Host "Removed: $p" -ForegroundColor Yellow
        }
    }
    Write-Host "Cleaned registry for: $name" -ForegroundColor Cyan
}

Write-Host "`nReinstalling packages via winget..." -ForegroundColor Green
$packages = @('GitHub.cli', 'Neovim.Neovim', 'Kitware.CMake', 'calibre.calibre')
foreach ($pkg in $packages) {
    Write-Host "Installing $pkg..." -ForegroundColor Cyan
    $result = winget install --id $pkg --silent --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
        Write-Host "[OK] $pkg" -ForegroundColor Green
    } else {
        Write-Host "[FAIL] $pkg (exit $LASTEXITCODE)" -ForegroundColor Red
        $result | Select-Object -Last 5 | Write-Host
    }
}
