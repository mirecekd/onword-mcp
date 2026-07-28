<#
.SYNOPSIS
    Emergency stop for onword-mcp: kills the launcher, the MCP server and the
    reverse SSH tunnel.

.DESCRIPTION
    onword-launch.ps1 normally cleans up after itself when Word is closed.
    Use this script when it did not (killed launcher, crash, reboot pending)
    and port 18347 is still taken.

    It first tries the PIDs recorded in %LOCALAPPDATA%\onword-mcp\launcher.json
    and then falls back to whatever process is still listening on the port.
#>
[CmdletBinding()]
param(
    # 0 = take ONWORD_PORT from onword.env / environment, else 18347.
    [int]$Port = 0
)

$ErrorActionPreference = "Continue"

if ($Port -le 0) {
    $Port = if ($env:ONWORD_PORT) { [int]$env:ONWORD_PORT } else { 18347 }
    $envFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "onword.env"
    if (-not $env:ONWORD_PORT -and (Test-Path $envFile)) {
        $match = Select-String -Path $envFile -Pattern '^\s*ONWORD_PORT\s*=\s*(\d+)' |
            Select-Object -First 1
        if ($match) { $Port = [int]$match.Matches[0].Groups[1].Value }
    }
}

$stateFile = Join-Path $env:LOCALAPPDATA "onword-mcp\launcher.json"
$killed = @()


function Stop-ByPid {
    param([int]$TargetPid, [string]$What)
    if (-not $TargetPid) { return }
    $proc = Get-Process -Id $TargetPid -ErrorAction SilentlyContinue
    if (-not $proc) { return }
    Write-Host "killing $What (pid $TargetPid, $($proc.ProcessName))"
    & taskkill.exe /PID $TargetPid /T /F 2>&1 | Out-Null
    $script:killed += $TargetPid
}

if (Test-Path $stateFile) {
    $state = Get-Content $stateFile -Raw | ConvertFrom-Json
    Stop-ByPid -TargetPid $state.ssh_pid -What "SSH tunnel"
    Stop-ByPid -TargetPid $state.server_pid -What "MCP server"
    Stop-ByPid -TargetPid $state.launcher_pid -What "launcher"
    Remove-Item $stateFile -ErrorAction SilentlyContinue
} else {
    Write-Host "no state file ($stateFile) - falling back to port scan"
}

# whatever still holds the port
$listeners = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
foreach ($l in $listeners) {
    if ($killed -notcontains $l.OwningProcess) {
        Stop-ByPid -TargetPid $l.OwningProcess -What "listener on port $Port"
    }
}

# leftover reverse tunnels for this port
Get-CimInstance Win32_Process -Filter "Name = 'ssh.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -and $_.CommandLine -match "${Port}:127\.0\.0\.1:$Port" } |
    ForEach-Object {
        if ($killed -notcontains $_.ProcessId) {
            Stop-ByPid -TargetPid $_.ProcessId -What "leftover ssh tunnel"
        }
    }

if ($killed.Count -eq 0) {
    Write-Host "nothing to stop - onword-mcp is not running"
} else {
    Write-Host "stopped $($killed.Count) process(es). Word itself was left untouched."
}
