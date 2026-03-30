New-Item -ItemType Directory -Path 'C:\Windows\Installer' -Force | Out-Null
icacls 'C:\Windows\Installer' /grant 'SYSTEM:(OI)(CI)F' /grant 'Administrators:(OI)(CI)F' /grant 'NETWORK SERVICE:(OI)(CI)F' /grant 'Users:(OI)(CI)RX'
Write-Host 'C:\Windows\Installer created successfully.' -ForegroundColor Green
