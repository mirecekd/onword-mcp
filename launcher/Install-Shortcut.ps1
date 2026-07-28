<#
.SYNOPSIS
    Creates the desktop shortcuts for onword-mcp (and the onword.env config).

.DESCRIPTION
    Creates on the Desktop:
      - "onword"          -> onword-launch.vbs (Word + MCP server + SSH tunnel)
      - "onword (update)" -> same, but with -NoCache (rebuild from GitHub)
      - "onword (stop)"   -> onword-stop.ps1   (only with -WithStop)
      - "onword (status)" -> onword-status.ps1 (only with -WithStatus)

    All settings (SSH host, port, source) live in onword.env next to the
    scripts - never in the shortcut and never in the repository. If
    onword.env does not exist it is created from onword.env.example; pass
    -SshTarget to fill in the host right away.

    The shortcuts point at the scripts in this directory, so keep the folder
    where it is (e.g. C:\tools\onword-mcp\launcher) or re-run this script
    after moving it.

    Because "onword" targets a .vbs run by wscript, no console window appears
    and the shortcut also accepts a dropped .docx file.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File Install-Shortcut.ps1

.EXAMPLE
    .\Install-Shortcut.ps1 -SshTarget dev.domain.com -Port 18347 -WithStop
#>
[CmdletBinding()]
param(
    # Where to put the shortcuts.
    [string]$Destination = [Environment]::GetFolderPath("Desktop"),

    [string]$Name = "onword",

    # Written into onword.env (not into the shortcut).
    [string]$SshTarget = "",
    [int]$Port = 0,
    [string]$Source = "",

    # Also create the "(update)" shortcut that forces a fresh uvx build.
    [switch]$WithUpdate = $true,

    # Also create the "(stop)" shortcut.
    [switch]$WithStop,

    # Also create the "(status)" shortcut (window stays open with the report).
    [switch]$WithStatus
)

$ErrorActionPreference = "Stop"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$launchVbs = Join-Path $here "onword-launch.vbs"
$stopPs1 = Join-Path $here "onword-stop.ps1"
$statusPs1 = Join-Path $here "onword-status.ps1"
$envFile = Join-Path $here "onword.env"
$envExample = Join-Path $here "onword.env.example"

if (-not (Test-Path $launchVbs)) { throw "not found: $launchVbs" }
if (-not (Test-Path $Destination)) { throw "destination does not exist: $Destination" }

# --- configuration file ----------------------------------------------------

if (-not (Test-Path $envFile)) {
    if (Test-Path $envExample) {
        Copy-Item $envExample $envFile
        Write-Host "created config: $envFile (from onword.env.example)"
    } else {
        New-Item -ItemType File -Path $envFile | Out-Null
        Write-Host "created empty config: $envFile"
    }
}

function Set-ConfValue {
    param([string]$Key, [string]$Value)
    if (-not $Value) { return }
    $lines = @(Get-Content $envFile -ErrorAction SilentlyContinue)
    $pattern = "^\s*#?\s*$Key\s*="
    if ($lines -match $pattern) {
        $lines = $lines | ForEach-Object {
            if ($_ -match $pattern) { "$Key=$Value" } else { $_ }
        }
    } else {
        $lines += "$Key=$Value"
    }
    Set-Content -Path $envFile -Value $lines -Encoding UTF8
    Write-Host "config: $Key=$Value"
}

Set-ConfValue -Key "ONWORD_SSH_TARGET" -Value $SshTarget
if ($Port -gt 0) { Set-ConfValue -Key "ONWORD_PORT" -Value "$Port" }
Set-ConfValue -Key "ONWORD_SOURCE" -Value $Source

# --- shortcuts -------------------------------------------------------------

# Word's own icon makes the shortcut recognizable; fall back to a shell icon.
$wordExe = (Get-Command winword.exe -ErrorAction SilentlyContinue).Source
if (-not $wordExe) {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
    )
    $wordExe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}
$iconLocation = if ($wordExe) { "$wordExe,0" } else { "$env:SystemRoot\System32\shell32.dll,13" }

$shell = New-Object -ComObject WScript.Shell

function New-Shortcut {
    param(
        [string]$LinkName,
        [string]$Target,
        [string]$Arguments,
        [string]$Description,
        [string]$Icon
    )
    $path = Join-Path $Destination "$LinkName.lnk"
    $sc = $shell.CreateShortcut($path)
    $sc.TargetPath = $Target
    $sc.Arguments = $Arguments
    $sc.WorkingDirectory = $here
    $sc.Description = $Description
    $sc.IconLocation = $Icon
    $sc.Save()
    Write-Host "created: $path"
}

$wscript = Join-Path $env:SystemRoot "System32\wscript.exe"
$powershell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

New-Shortcut -LinkName $Name -Target $wscript `
    -Arguments "`"$launchVbs`"" `
    -Description "Start Word + onword-mcp server + reverse SSH tunnel" `
    -Icon $iconLocation

if ($WithUpdate) {
    New-Shortcut -LinkName "$Name (update)" -Target $wscript `
        -Arguments "`"$launchVbs`" -NoCache" `
        -Description "Same as onword, but rebuilds the server from GitHub (slow)" `
        -Icon "$env:SystemRoot\System32\shell32.dll,238"
}

if ($WithStop) {
    New-Shortcut -LinkName "$Name (stop)" -Target $powershell `
        -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$stopPs1`"" `
        -Description "Kill the onword-mcp server and the SSH tunnel" `
        -Icon "$env:SystemRoot\System32\shell32.dll,131"
}

if ($WithStatus) {
    # -NoExit so the report stays readable after a double-click.
    New-Shortcut -LinkName "$Name (status)" -Target $powershell `
        -Arguments "-NoProfile -NoExit -ExecutionPolicy Bypass -File `"$statusPs1`"" `
        -Description "Show whether Word, the onword-mcp server and the tunnel are up" `
        -Icon "$env:SystemRoot\System32\shell32.dll,23"
}

Write-Host ""
Write-Host "Done. Review $envFile, then double-click '$Name' on the Desktop."
Write-Host "Logs: $env:LOCALAPPDATA\onword-mcp\logs"
