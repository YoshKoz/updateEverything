

function Write-Console
{
    <#
    .SYNOPSIS
        Console output with colour, without Write-Host.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][AllowEmptyString()][string]$Message = '',
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor,
        [switch]$NoNewline
    )

    $fg = [Console]::ForegroundColor
    $bg = [Console]::BackgroundColor
    try
    {
        if ($PSBoundParameters.ContainsKey('ForegroundColor')) { [Console]::ForegroundColor = $ForegroundColor }
        if ($PSBoundParameters.ContainsKey('BackgroundColor')) { [Console]::BackgroundColor = $BackgroundColor }
        if ($NoNewline) { [Console]::Write($Message) } else { [Console]::WriteLine($Message) }
    }
    finally
    {
        [Console]::ForegroundColor = $fg
        [Console]::BackgroundColor = $bg
    }
}
$exe = "C:\Development\updateEverything\rewrites\rust\target\release\updateeverything.exe"
$logFile = "C:\Development\updateEverything\rewrites\rust\run-$(Get-Date -Format yyyyMMdd-HHmmss).log"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
& $exe --jobs 4 *>&1 | Tee-Object -FilePath $logFile
$sw.Stop()
Write-Console "ELAPSED: $($sw.Elapsed)" -ForegroundColor Cyan
