# onword-mcp launcher (Windows)

One desktop icon that starts everything needed for remote LLM editing of the
document open in Word:

1. **Word** (if not already running; a `.docx` can be passed/dropped)
2. **the MCP server** - `uvx --from git+https://github.com/mirecekd/onword-mcp python -m onword_mcp --transport streamable-http --port 18347 --host 127.0.0.1`
3. **the reverse SSH tunnel** - `ssh -N -R 18347:127.0.0.1:18347 <your-host>`

The launcher then supervises the pieces: it reconnects the SSH tunnel whenever
it drops, and kills both the server and the tunnel as soon as Word is closed,
so the next start always finds a free port.

## Install

Copy the repository (or just this `launcher` folder plus nothing else - the
server is fetched from GitHub by `uvx`) to the Windows machine, e.g.
`C:\tools\onword-mcp`, then:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File C:\tools\onword-mcp\launcher\Install-Shortcut.ps1 -SshTarget dev.domain.com -WithStop -WithStatus
```

This creates `onword.env` (from `onword.env.example`) with your SSH host, and
icons on the Desktop:

| Icon | What it does |
|---|---|
| **onword** | starts Word + server + tunnel (cached build, starts in seconds) |
| **onword (update)** | same, but `--no-cache` - re-downloads and rebuilds from GitHub (slow, needs network); use after pushing new commits |
| **onword (stop)** | kills a leftover server/tunnel if cleanup did not happen |
| **onword (status)** | read-only check: is Word up, is the server serving, is the tunnel alive (`-WithStatus`) |

Deleting the icons does not stop anything and does not uninstall anything -
they are plain `.lnk` files pointing at the scripts in this folder. Re-run
`Install-Shortcut.ps1` to get them back (existing `onword.env` is kept).

Then just double-click **onword**. Nothing flashes on the screen - the whole
chain runs hidden and logs into `%LOCALAPPDATA%\onword-mcp\logs`.

You can also drag a `.docx` onto the **onword** icon; it is opened in Word.

## Configuration: onword.env

Host names and ports are **not** in the scripts or in the shortcuts - they live
in `onword.env` next to the scripts (or in
`%LOCALAPPDATA%\onword-mcp\onword.env`). That file is gitignored; the template
is `onword.env.example`:

```ini
ONWORD_SSH_TARGET=dev.domain.com
ONWORD_PORT=18347
#ONWORD_REMOTE_PORT=
#ONWORD_SOURCE=git+https://github.com/mirecekd/onword-mcp
#ONWORD_NO_TUNNEL=0
```

Precedence: command-line parameter > real environment variable > `onword.env` >
built-in default. If `ONWORD_SSH_TARGET` is empty, the launcher just runs
locally without a tunnel and says so in the log.

`ONWORD_SSH_TARGET` can be a `Host` alias from `~/.ssh/config`, which is the
clean way to pin a user, port or identity file:

```
Host onword-relay
    HostName dev.domain.com
    User me
    IdentityFile ~/.ssh/id_ed25519
```

## Prerequisites

- `uv` on PATH: `winget install astral-sh.uv`
- OpenSSH client on PATH (built into Windows 10/11)
- **passwordless SSH** to the target host - the tunnel runs hidden with
  `BatchMode=yes`, so a password prompt nobody can see would just fail.
  Set up a key (`ssh-keygen`, `type $env:USERPROFILE\.ssh\id_ed25519.pub` ->
  append to `~/.ssh/authorized_keys` on the server), or add a passphrase-
  protected key to `ssh-agent` (`Set-Service ssh-agent -StartupType Automatic;
  Start-Service ssh-agent; ssh-add`).

## Manual use

```powershell
# everything from onword.env
.\onword-launch.ps1

# open a document, fresh build, output visible in the console
.\onword-launch.ps1 -Document C:\docs\report.docx -NoCache

# local only, no tunnel
.\onword-launch.ps1 -NoTunnel

# server + tunnel without touching Word (runs until Ctrl+C / stop script)
.\onword-launch.ps1 -NoWord

# override config values ad hoc
.\onword-launch.ps1 -SshTarget other.domain.com -Port 18400 -Source C:\tools\onword-mcp

# use a different config file
.\onword-launch.ps1 -EnvFile C:\secrets\onword-work.env
```

Check what is running:

```powershell
.\onword-status.ps1

# same plus the tail of launcher/server/ssh logs
.\onword-status.ps1 -Log
```

Output is one line per component (`[UP]` / `[DOWN]`) for Word, the MCP server
(with the PID holding the port and the endpoint URL), the SSH tunnel and the
launcher, plus a warning about a stale `launcher.json`. Exit code is 0 when the
server serves, so it can be used in scripts:
`.\onword-status.ps1 > $null; if ($LASTEXITCODE) { .\onword-launch.ps1 }`.

Stop a leftover instance:

```powershell
.\onword-stop.ps1
```

## Auto-start with Word (optional)

`AutoExec.bas` is a small VBA module for `Normal.dotm` that runs the launcher
whenever Word starts (including documents opened from Teams/SharePoint):

1. Word -> `Alt+F11` -> File -> Import File... -> `AutoExec.bas`
2. Adjust `LAUNCHER_PATH` inside the module
3. Ctrl+S to save `Normal`, restart Word

The launcher is idempotent - if the port already serves, it exits immediately,
so a second Word window does not start a second server.

Caveat: corporate macro policies (or the AV heuristic for `Shell` inside
`AutoExec`) often block macros in `Normal.dotm`. The desktop icon is the
robust path; treat the macro as a convenience on top.

## Notes / design

- **Startup order does not matter.** The server never caches COM objects; it
  looks up the running Word instance on every tool call, so it happily starts
  before Word finishes loading.
- **Bind address is `127.0.0.1`,** not `0.0.0.0` - the SSH tunnel connects
  locally, and binding to all interfaces only triggers a Windows Firewall
  prompt and exposes the port to the LAN.
- **`-R` binds `127.0.0.1:<port>` on the remote host** unless the server has
  `GatewayPorts yes`, which is what you want for a local client or a reverse
  proxy there.
- **Tunnel resilience:** `ExitOnForwardFailure=yes` so a stale remote port is
  detected instead of silently ignored, plus `ServerAliveInterval=30` /
  `ServerAliveCountMax=3` so a dead link is noticed within ~90 s and
  reconnected by the supervision loop (throttled to avoid reconnect storms).
- **Cleanup** uses `taskkill /T` because `uvx` runs the real server in a child
  python process.

## Logs

`%LOCALAPPDATA%\onword-mcp\logs`

| File | Content |
|---|---|
| `launcher.log` | config used, launcher decisions, restarts, shutdown |
| `server.out.log` / `server.err.log` | MCP server output (uvx build errors land here) |
| `ssh.err.log` | SSH tunnel diagnostics |

`%LOCALAPPDATA%\onword-mcp\launcher.json` holds the current PIDs and is
removed on clean shutdown.
