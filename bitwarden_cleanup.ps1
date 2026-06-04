#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Find and merge duplicate Bitwarden vault entries.
.PARAMETER DryRun
    Show duplicates without making changes.
.PARAMETER AutoMerge
    Automatically merge without prompt.
.PARAMETER Session
    Session key (or set $env:BW_SESSION).
#>
param(
    [switch]$DryRun,
    [switch]$AutoMerge,
    [string]$Session = $env:BW_SESSION
)

$ErrorActionPreference = 'Stop'

function Run-Bw {
    param([Parameter(ValueFromRemainingArguments)][string[]]$Args)
    $allArgs = @($Args)
    if ($script:SessionKey) { $allArgs += @("--session", $script:SessionKey) }
    $out = & bw @allArgs 2>$null
    if ($LASTEXITCODE -ne 0) { throw "bw $($Args[0]) failed (exit $LASTEXITCODE)" }
    return $out
}

$script:SessionKey = if ($Session) { $Session } else { $env:BW_SESSION }
if (-not $script:SessionKey) { Write-Host "[X] No session. Run 'bw unlock' and set `$env:BW_SESSION" -ForegroundColor Red; exit 1 }

Write-Host "`n=== Bitwarden Vault Cleanup ===" -ForegroundColor Magenta

$status = Run-Bw status | ConvertFrom-Json
if ($status.status -ne 'unlocked') { Write-Host "[X] Vault not unlocked: $($status.status)" -ForegroundColor Red; exit 1 }
Write-Host "  [OK] Logged in as $($status.userEmail)" -ForegroundColor Green

Write-Host "`nSyncing vault..." -ForegroundColor White
Run-Bw sync | Out-Null
Write-Host "  [OK] Synced" -ForegroundColor Green

$backupFile = "bw_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
Write-Host "`nExporting backup to $backupFile..." -ForegroundColor White
Run-Bw export --format json --raw | Set-Content -Path $backupFile -Encoding utf8
Write-Host "  [OK] Backup saved" -ForegroundColor Green

Write-Host "`nLoading items..." -ForegroundColor White
$items = Run-Bw list items | ConvertFrom-Json
$logins = $items | Where-Object { $_.type -eq 1 }
Write-Host "  [OK] $($items.Count) total, $($logins.Count) logins" -ForegroundColor Green

Write-Host "`nFinding duplicates..." -ForegroundColor White

function Get-DupKey($item) {
    $name = ($item.name ?? '').Trim().ToLower()
    $user = ($item.login.username ?? '').Trim().ToLower()
    $uris = (($item.login.uris | ForEach-Object { $_.uri }) -join '|') -replace 'https?://', '' -replace 'www\.', '' -replace '/$', ''
    return "$name|$user|$uris"
}

$groups = @{}
foreach ($item in $logins) {
    $key = Get-DupKey $item
    if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
    $groups[$key] += $item
}

$dups = $groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 } | Sort-Object { $_.Value[0].name }
$dupCount = @($dups).Count

if ($dupCount -eq 0) {
    Write-Host "`n  [OK] No duplicates found!" -ForegroundColor Green
    exit 0
}

$totalDupItems = ($dups | ForEach-Object { $_.Value.Count } | Measure-Object -Sum).Sum
Write-Host "  [!] Found $dupCount duplicate groups ($totalDupItems items)" -ForegroundColor Yellow

Write-Host "`nDuplicate groups:" -ForegroundColor White
$groupNum = 0
foreach ($group in $dups) {
    $groupNum++
    $sorted = $group.Value | Sort-Object { $_.revisionDate } -Descending
    $keep = $sorted[0]
    
    Write-Host "`n  --- Group $groupNum ---" -ForegroundColor Yellow
    Write-Host "  Name: $($keep.name)" -ForegroundColor Cyan
    Write-Host "  User: $($keep.login.username)" -ForegroundColor Cyan
    if ($keep.login.uris) { Write-Host "  URL:  $($keep.login.uris[0].uri)" -ForegroundColor Cyan }
    
    foreach ($item in $sorted) {
        $rev = [datetime]$item.revisionDate
        $pwd = if ($item.login.password) { "pwd" } else { "no-pwd" }
        $totp = if ($item.login.totp) { "+totp" } else { "" }
        $notes = if ($item.notes) { "+notes" } else { "" }
        $flds = if ($item.fields) { "+$($item.fields.Count)flds" } else { "" }
        $tag = if ($item.id -eq $keep.id) { "[KEEP]" } else { "[DEL]" }
        $color = if ($tag -eq '[KEEP]') { 'Green' } else { 'DarkGray' }
        Write-Host "    $($item.id.Substring(0,8)) $($rev.ToString('yyyy-MM-dd')) $pwd$totp$notes$flds $tag" -ForegroundColor $color
    }
}

if ($DryRun) {
    Write-Host "`n[Dry Run] No changes made." -ForegroundColor Yellow
    exit 0
}

if (-not $AutoMerge) {
    Write-Host "`nProceed with merge? (y/N): " -ForegroundColor White -NoNewline
    if ((Read-Host) -notin 'y','Y') { Write-Host "Cancelled." -ForegroundColor Yellow; exit 0 }
}

Write-Host "`nMerging..." -ForegroundColor White
$merged = 0; $deleted = 0

foreach ($group in $dups) {
    $sorted = $group.Value | Sort-Object { $_.revisionDate } -Descending
    $keep = $sorted[0]
    $remove = $sorted | Select-Object -Skip 1
    $changed = $false

    foreach ($dup in $remove) {
        if ($dup.notes -and -not $keep.notes) { $keep.notes = $dup.notes; $changed = $true }
        elseif ($dup.notes -and $keep.notes -and $dup.notes -ne $keep.notes) {
            $keep.notes = "$($keep.notes)`n---merged---`n$($dup.notes)"; $changed = $true
        }
        if ($dup.fields) {
            foreach ($f in $dup.fields) {
                if (-not ($keep.fields | Where-Object { $_.name -eq $f.name })) {
                    if (-not $keep.fields) { $keep.fields = @() }
                    $keep.fields += $f; $changed = $true
                }
            }
        }
        if (-not $keep.login.password -and $dup.login.password) { $keep.login.password = $dup.login.password; $changed = $true }
        if (-not $keep.login.totp -and $dup.login.totp) { $keep.login.totp = $dup.login.totp; $changed = $true }
        if ($dup.login.uris) {
            foreach ($u in $dup.login.uris) {
                if (-not ($keep.login.uris | Where-Object { $_.uri -eq $u.uri })) {
                    if (-not $keep.login.uris) { $keep.login.uris = @() }
                    $keep.login.uris += $u; $changed = $true
                }
            }
        }
    }

    if ($changed) {
        $json = $keep | ConvertTo-Json -Depth 10 -Compress
        $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
        Run-Bw edit item $keep.id $encoded | Out-Null
        Write-Host "  [OK] Updated: $($keep.name)" -ForegroundColor Green
        $merged++
    }

    foreach ($dup in $remove) {
        try {
            Run-Bw delete item $dup.id | Out-Null
            Write-Host "  [OK] Deleted: $($dup.name) ($($dup.id.Substring(0,8)))" -ForegroundColor Green
            $deleted++
        } catch {
            Write-Host "  [X] Delete failed $($dup.id): $_" -ForegroundColor Red
        }
    }
}

Write-Host "`n=== Done ===" -ForegroundColor Magenta
Write-Host "  Merged: $merged | Deleted: $deleted" -ForegroundColor Green
Write-Host "  Backup: $backupFile" -ForegroundColor Cyan
