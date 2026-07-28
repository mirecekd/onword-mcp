<#
.SYNOPSIS
    Status check for onword-mcp: is Word running, is the MCP server serving,
    is the reverse SSH tunnel up?

.DESCRIPTION
    Read-only counterpart of onword-stop.ps1. It reports:

      - Word            (WINWORD.EXE processes)
      - MCP server      (TCP port serving on 127.0.0.1 + owning process)
      - SSH tunnel      (ssh.exe with -R <port>:127.0.0.1:<port>)
      - launcher        (PIDs recorded in %LOCALAPPDATA%\onword-mcp\launcher.json)

    Exit code 0 = server is serving, 1 = it is not (usable from scripts).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File onword-status.ps1

.EXAMPLE
    # only the last lines of the logs, when something looks wrong
    .\onword-status.ps1 -Log
#>
[CmdletBinding()]
param(
    # 0 = take ONWORD_PORT from onword.env / environment, else 18347.
    [int]$Port = 0,

    # Also print the tail of the launcher/server/ssh logs.
    [switch]$Log,

    # How many lines of each log to show with -Log.
    [int]$Lines = 15
)

$ErrorActionPreference = "Continue"

if ($Port -le 0) {
    if ($env:ONWORD_PORT) { $Port = [int]$env:ONWORD_PORT } else { $Port = 18347 }
    $envFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "onword.env"
    if (-not $env:ONWORD_PORT -and (Test-Path $envFile)) {
        $match = Select-String -Path $envFile -Pattern '^\s*ONWORD_PORT\s*=\s*(\d+)' |
            Select-Object -First 1
        if ($match) { $Port = [int]$match.Matches[0].Groups[1].Value }
    }
}

$stateDir = Join-Path $env:LOCALAPPDATA "onword-mcp"
$stateFile = Join-Path $stateDir "launcher.json"
$logDir = Join-Path $stateDir "logs"

function Write-Status {
    param([string]$What, [bool]$Ok, [string]$Detail = "")
    $tag = if ($Ok) { "[UP]  " } else { "[DOWN]" }
    Write-Host ("{0} {1,-12} {2}" -f $tag, $What, $Detail)
}

function Test-Port {
    param([int]$TcpPort)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $client.Connect("127.0.0.1", $TcpPort)
        return $true
    } catch {
        return $false
    } finally {
        $client.Dispose()
    }
}

function Get-ProcDescription {
    param([int]$TargetPid)
    $proc = Get-Process -Id $TargetPid -ErrorAction SilentlyContinue
    if (-not $proc) { return "pid $TargetPid (gone)" }
    return "pid $TargetPid ($($proc.ProcessName))"
}

Write-Host "onword-mcp status (port $Port)"
Write-Host ""

# --- Word ------------------------------------------------------------------

$word = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
$wordInfo = if ($word.Count -gt 0) {
    "$($word.Count) process(es): $(($word | ForEach-Object { $_.Id }) -join ', ')"
} else {
    "not running - the MCP tools will fail until a document is open"
}
Write-Status -What "Word" -Ok ($word.Count -gt 0) -Detail $wordInfo

# --- MCP server ------------------------------------------------------------

$serving = Test-Port -TcpPort $Port
$listeners = @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)
$listenerInfo = if ($listeners.Count -gt 0) {
    ($listeners | ForEach-Object { Get-ProcDescription -TargetPid $_.OwningProcess } | Select-Object -Unique) -join ', '
} else { "no listener on 127.0.0.1:$Port" }
Write-Status -What "MCP server" -Ok $serving -Detail $listenerInfo
if ($serving) {
    Write-Host ("       endpoint: http://127.0.0.1:{0}/mcp" -f $Port)
}

# --- SSH tunnel ------------------------------------------------------------

$tunnels = @(
    Get-CimInstance Win32_Process -Filter "Name = 'ssh.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -and $_.CommandLine -match "${Port}:127\.0\.0\.1:$Port" }
)
$tunnelInfo = if ($tunnels.Count -gt 0) {
    ($tunnels | ForEach-Object { "pid $($_.ProcessId)" }) -join ', '
} else {
    "no reverse tunnel for this port (fine with ONWORD_NO_TUNNEL=1)"
}
Write-Status -What "SSH tunnel" -Ok ($tunnels.Count -gt 0) -Detail $tunnelInfo
foreach ($t in $tunnels) { Write-Host ("       $($t.CommandLine)") }

# --- launcher / state file -------------------------------------------------

if (Test-Path $stateFile) {
    $state = Get-Content $stateFile -Raw | ConvertFrom-Json
    $launcherAlive = [bool](Get-Process -Id $state.launcher_pid -ErrorAction SilentlyContinue)
    $launcherInfo = "$(Get-ProcDescription -TargetPid $state.launcher_pid), started $($state.started)"
    Write-Status -What "launcher" -Ok $launcherAlive -Detail $launcherInfo
    Write-Host ("       state: server $(Get-ProcDescription -TargetPid $state.server_pid), ssh $(Get-ProcDescription -TargetPid $state.ssh_pid)")
    if (-not $launcherAlive) {
        Write-Host "       stale state file - run onword-stop.ps1 to clean up"
    }
} else {
    Write-Status -What "launcher" -Ok $false -Detail "no state file ($stateFile) - clean shutdown or never started"
}

# --- logs ------------------------------------------------------------------

if ($Log) {
    foreach ($name in @("launcher.log", "server.err.log", "server.out.log", "ssh.err.log")) {
        $path = Join-Path $logDir $name
        Write-Host ""
        if (Test-Path $path) {
            Write-Host "--- $path (last $Lines lines) ---"
            Get-Content $path -Tail $Lines
        } else {
            Write-Host "--- $path (missing) ---"
        }
    }
} else {
    Write-Host ""
    Write-Host "logs: $logDir   (re-run with -Log to tail them)"
}

if ($serving) { exit 0 } else { exit 1 }
