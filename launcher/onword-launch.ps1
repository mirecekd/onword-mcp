<#
.SYNOPSIS
    One-click launcher for onword-mcp: starts Word, the MCP server and the
    reverse SSH tunnel, keeps the tunnel alive, and cleans everything up as
    soon as Word is closed.

.DESCRIPTION
    Order of startup does not matter - the MCP server attaches to the running
    Word instance lazily on every tool call - so this script just brings all
    three pieces up and then supervises them:

      1. Word            (only if not already running; optional document)
      2. MCP server      uvx --from <source> python -m onword_mcp ...
      3. SSH tunnel      ssh -N -R <port>:127.0.0.1:<port> <target>

    While Word is running the script re-establishes the SSH tunnel whenever it
    drops (network change, laptop sleep, server restart). When Word exits, the
    server and the tunnel are terminated so the next start finds a free port.

    Everything is logged to %LOCALAPPDATA%\onword-mcp\logs.

    Configuration (SSH host, port, source) comes from onword.env - see
    onword.env.example. Precedence: parameter > environment variable >
    onword.env > built-in default.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File onword-launch.ps1

.EXAMPLE
    # open a document and refresh the server build from GitHub
    .\onword-launch.ps1 -Document C:\docs\report.docx -NoCache
#>
[CmdletBinding()]
param(
    # Document to open in Word (optional).
    [string]$Document,

    # Local port the MCP server listens on. 0 = from config (ONWORD_PORT).
    [int]$Port = 0,

    # SSH host for the reverse tunnel (ONWORD_SSH_TARGET in onword.env).
    # Empty and no config value, or -NoTunnel, disables the tunnel.
    [string]$SshTarget = "",

    # Remote port to bind on the SSH host. 0 = same as -Port.
    [int]$RemotePort = 0,

    # What uvx should build the server from: a git URL, a local directory
    # or a PyPI package name (ONWORD_SOURCE).
    [string]$Source = "",

    # Explicit config file; default is onword.env next to this script, then
    # %LOCALAPPDATA%\onword-mcp\onword.env.
    [string]$EnvFile = "",


    # Force a fresh download/build (slow, needs network). Use after pushing
    # new commits; otherwise the cached build starts instantly.
    [switch]$NoCache,

    # Do not start/require Word (server then runs until stopped manually).
    [switch]$NoWord,

    # Do not start the SSH tunnel (local-only usage).
    [switch]$NoTunnel
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- configuration: onword.env --------------------------------------------
# Keeps host names out of the repository. Parameters and real environment
# variables win over the file, the file wins over the defaults below.

function Import-EnvFile {
    param([string]$Path)
    $map = @{}
    foreach ($line in Get-Content -Path $Path -Encoding UTF8) {
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith("#")) { continue }
        $eq = $t.IndexOf("=")
        if ($eq -lt 1) { continue }
        $key = $t.Substring(0, $eq).Trim()
        $val = $t.Substring($eq + 1).Trim().Trim('"').Trim("'")
        $map[$key] = $val
    }
    return $map
}

$envCandidates = @()
if ($EnvFile) { $envCandidates += $EnvFile }
$envCandidates += (Join-Path $scriptDir "onword.env")
$envCandidates += (Join-Path $env:LOCALAPPDATA "onword-mcp\onword.env")

$conf = @{}
$usedEnvFile = $null
foreach ($candidate in $envCandidates) {
    if ($candidate -and (Test-Path $candidate)) {
        $conf = Import-EnvFile -Path $candidate
        $usedEnvFile = $candidate
        break
    }
}

function Get-Conf {
    param([string]$Key, [string]$Default = "")
    $fromEnv = [Environment]::GetEnvironmentVariable($Key)
    if ($fromEnv) { return $fromEnv }
    if ($conf.ContainsKey($Key) -and $conf[$Key]) { return $conf[$Key] }
    return $Default
}

if (-not $SshTarget) { $SshTarget = Get-Conf "ONWORD_SSH_TARGET" }
if (-not $Source) { $Source = Get-Conf "ONWORD_SOURCE" "git+https://github.com/mirecekd/onword-mcp" }
if ($Port -le 0) { $Port = [int](Get-Conf "ONWORD_PORT" "18347") }
if ($RemotePort -le 0) { $RemotePort = [int](Get-Conf "ONWORD_REMOTE_PORT" "0") }
if ((Get-Conf "ONWORD_NO_TUNNEL") -eq "1") { $NoTunnel = $true }

if ($RemotePort -le 0) { $RemotePort = $Port }

$stateDir = Join-Path $env:LOCALAPPDATA "onword-mcp"

$logDir = Join-Path $stateDir "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$launcherLog = Join-Path $logDir "launcher.log"
$stateFile = Join-Path $stateDir "launcher.json"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $launcherLog -Value $line
    Write-Host $line
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

function Get-WordProcess {
    Get-Process -Name WINWORD -ErrorAction SilentlyContinue
}

function Stop-Tree {
    param([System.Diagnostics.Process]$Process, [string]$What)
    if (-not $Process) { return }
    if ($Process.HasExited) { return }
    Write-Log "stopping $What (pid $($Process.Id))"
    try {
        # /T also kills the python child that uvx spawned
        & taskkill.exe /PID $Process.Id /T /F 2>&1 | Out-Null
    } catch {
        try { $Process.Kill() } catch { }
    }
}

function Save-State {
    param($ServerPid, $SshPid)
    $state = [ordered]@{
        launcher_pid = $PID
        server_pid   = $ServerPid
        ssh_pid      = $SshPid
        port         = $Port
        started      = (Get-Date -Format "s")
    }
    ($state | ConvertTo-Json) | Set-Content -Path $stateFile -Encoding UTF8
}

Write-Log "----- launcher start (port $Port, target '$SshTarget') -----"
if ($usedEnvFile) {
    Write-Log "config: $usedEnvFile"
} else {
    Write-Log "no onword.env found (looked in: $($envCandidates -join '; ')) - using defaults" "WARN"
}
if (-not $SshTarget -and -not $NoTunnel) {
    Write-Log "ONWORD_SSH_TARGET is not set - running without a reverse tunnel. Copy onword.env.example to onword.env to configure it." "WARN"
}


# --- sanity checks ---------------------------------------------------------

if (-not (Get-Command uvx -ErrorAction SilentlyContinue)) {
    Write-Log "uvx not found in PATH. Install uv: winget install astral-sh.uv" "ERROR"
    exit 1
}

if (Test-Port -TcpPort $Port) {
    Write-Log "port $Port is already serving - another onword-mcp instance is running. Nothing to do." "WARN"
    exit 0
}

# --- 1) Word --------------------------------------------------------------

if (-not $NoWord) {
    if (Get-WordProcess) {
        Write-Log "Word already running"
        if ($Document) {
            Write-Log "opening document in the running Word: $Document"
            Start-Process -FilePath $Document
        }
    } else {
        if ($Document) {
            Write-Log "starting Word with document: $Document"
            Start-Process -FilePath "winword.exe" -ArgumentList "`"$Document`""
        } else {
            Write-Log "starting Word"
            Start-Process -FilePath "winword.exe"
        }
    }
}

# --- 2) MCP server --------------------------------------------------------

$serverArgs = @()
if ($NoCache) { $serverArgs += "--no-cache" }
$serverArgs += @(
    "--from", $Source,
    "python", "-m", "onword_mcp",
    "--transport", "streamable-http",
    "--port", "$Port",
    "--host", "127.0.0.1"
)

Write-Log "starting MCP server: uvx $($serverArgs -join ' ')"
$server = Start-Process -FilePath "uvx" -ArgumentList $serverArgs -PassThru `
    -WindowStyle Hidden `
    -RedirectStandardOutput (Join-Path $logDir "server.out.log") `
    -RedirectStandardError (Join-Path $logDir "server.err.log")

Save-State -ServerPid $server.Id -SshPid $null

$ssh = $null
try {
    # wait for the server to come up (a cold uvx build can take a while)
    $timeout = if ($NoCache) { 300 } else { 120 }
    $waited = 0
    while (-not (Test-Port -TcpPort $Port)) {
        if ($server.HasExited) {
            Write-Log "MCP server exited with code $($server.ExitCode) - see $logDir\server.err.log" "ERROR"
            exit 1
        }
        if ($waited -ge $timeout) {
            Write-Log "MCP server did not start listening on port $Port within ${timeout}s" "ERROR"
            exit 1
        }
        Start-Sleep -Seconds 2
        $waited += 2
    }
    Write-Log "MCP server listening on http://127.0.0.1:$Port/mcp"

    # --- 3) supervise: keep the tunnel up while Word is running ------------

    if (-not $NoWord) {
        # Word may still be initializing
        $waited = 0
        while (-not (Get-WordProcess) -and $waited -lt 60) {
            Start-Sleep -Seconds 2
            $waited += 2
        }
        if (-not (Get-WordProcess)) {
            Write-Log "Word process did not appear - shutting down again" "ERROR"
            exit 1
        }
    }

    $useTunnel = (-not $NoTunnel) -and $SshTarget
    if (-not $useTunnel) { Write-Log "SSH tunnel disabled" }

    while ($true) {
        if (-not $NoWord -and -not (Get-WordProcess)) {
            Write-Log "Word closed - shutting down"
            break
        }
        if ($server.HasExited) {
            Write-Log "MCP server exited with code $($server.ExitCode) - shutting down" "ERROR"
            break
        }

        if ($useTunnel -and (-not $ssh -or $ssh.HasExited)) {
            if ($ssh -and $ssh.HasExited) {
                Write-Log "SSH tunnel died (exit $($ssh.ExitCode)) - reconnecting" "WARN"
            }
            $sshArgs = @(
                "-N", "-T",
                "-o", "ExitOnForwardFailure=yes",
                "-o", "ServerAliveInterval=30",
                "-o", "ServerAliveCountMax=3",
                "-o", "BatchMode=yes",
                "-o", "StrictHostKeyChecking=accept-new",
                "-R", "${RemotePort}:127.0.0.1:$Port",
                $SshTarget
            )
            Write-Log "starting SSH tunnel: ssh $($sshArgs -join ' ')"
            $ssh = Start-Process -FilePath "ssh" -ArgumentList $sshArgs -PassThru `
                -WindowStyle Hidden `
                -RedirectStandardError (Join-Path $logDir "ssh.err.log")
            Save-State -ServerPid $server.Id -SshPid $ssh.Id
            Start-Sleep -Seconds 5   # throttle reconnect storms
        }

        Start-Sleep -Seconds 5
    }
} finally {
    Stop-Tree -Process $ssh -What "SSH tunnel"
    Stop-Tree -Process $server -What "MCP server"
    Remove-Item -Path $stateFile -ErrorAction SilentlyContinue
    Write-Log "----- launcher stop -----"
}
