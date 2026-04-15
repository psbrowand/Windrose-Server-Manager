# Windrose Server Manager

A local GUI application for running and managing a [Windrose](https://store.steampowered.com/app/3041230/Windrose/) dedicated server on Windows.

![Dashboard](assets/screenshot-dashboard.png)

---

## Features

- **One-click Start / Stop / Restart** with configurable countdown warning before restart
- **Live dashboard** — CPU usage, RAM, player count, uptime, and connected player list
- **Live log viewer** — color-coded, filterable (All / Players / Warnings / Errors) with auto-scroll
- **Console command input** — send commands directly to the server process (Save World, List Players, custom commands)
- **Config editor** — edit server name, max players, password, and all world difficulty settings (preset or custom sliders) without touching JSON files
- **One-click world backup** — zips your save data to a timestamped archive
- **Scheduled daily restart** — set a time; manager restarts the server automatically
- **Auto-restart on crash** — watchdog detects unexpected exits and relaunches automatically
- **Player history** — persistent log of who joined and left
- **Invite code share** — copies a ready-to-send message to clipboard
- **Install wizard** — auto-detects Windrose in your Steam library and installs the dedicated server with one click

---

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 (built into Windows — no install needed)
- .NET Framework 4.5 or later (pre-installed on Windows 10+)
- **Windrose** owned and installed via Steam (App ID 3041230)

> The dedicated server files are bundled inside the Windrose game install. You do not need a separate dedicated server download.

---

## Quick Start

### First time (new machine)

1. **Install Windrose** from Steam and let it fully download.
2. **Download this repository** — click *Code > Download ZIP* on GitHub and extract it, or clone it:
   ```
   git clone https://github.com/psbrowand/Windrose-Server-Manager.git "C:\Game-Servers\Windrose"
   ```
3. **Run `Launch.vbs`** — double-click it. The app opens with no terminal windows. The app opens to the **Install** tab automatically.
4. Click **Auto-Detect** to find your Windrose Steam installation.
5. Click **Install Server** — the server files (~2.8 GB) are copied to the manager folder.
6. Switch to the **Dashboard** tab and click **Start**.

### Already have server files

If you already set up a Windrose server in this folder, just run `Launch.vbs` and you are ready to go.

---

## Tab Guide

| Tab | What it does |
|---|---|
| **Dashboard** | Live stats, player list, auto-restart toggle |
| **Config** | Edit `ServerDescription.json` and `WorldDescription.json` via form fields and sliders |
| **Log** | Live-tailing server log with color coding and filters |
| **Console** | Send console commands to the running server |
| **Tools** | Backup, scheduled restart, restart countdown, player history |
| **Install** | Detect and copy server files from your Steam installation |

---

## Console Commands

From the **Console** tab you can type any Unreal Engine console command and send it to the server. Useful built-in commands:

| Button / Command | Effect |
|---|---|
| Save World | Force-saves the world (`SaveWorld`) |
| List Players | Prints connected players to the log (`listplayers`) |
| Server Info | Prints frame/unit stats (`stat unit`) |
| Quit Server | Gracefully shuts down the server (`quit`) |
| `kick <name>` | Kicks a player by name |
| `servertravel <map>` | Travels to a different map |

> Console commands are sent via stdin to the server process. Not all commands may be available depending on the server build.

---

## Config Files

The manager edits these files on your behalf (always stops the server first):

| File | Location | Purpose |
|---|---|---|
| `ServerDescription.json` | `R5\` | Server name, max players, password, invite code |
| `WorldDescription.json` | `R5\Saved\SaveProfiles\...\Worlds\<id>\` | Difficulty preset and all gameplay multipliers |

> The server rewrites these files on exit. Always use the **Stop** button (not killing the process) to ensure a clean save.

---

## Backups

Click **Backup Now** in the **Tools** tab to create a timestamped `.zip` of your entire world save. Backups are stored in the `Backups\` folder next to this README.

Restore a backup by stopping the server, extracting the zip over `R5\Saved\SaveProfiles\`, and restarting.

---

## Port Forwarding

For players outside your local network to connect, forward these ports on your router:

| Protocol | Port |
|---|---|
| UDP | 7777 |
| UDP | 7778 |

Players join via **Play > Connect to Server** in Windrose, using the invite code shown in the manager header.

---

## Scheduled Restart

In the **Tools** tab, enable the scheduled restart and enter a time in `HH:mm` 24-hour format (e.g. `04:00`). The manager will automatically restart the server at that time once per day, respecting the restart countdown warning you configured.

> The manager window must be open for scheduled restarts to fire.

---

## Folder Structure

```
Windrose-Server-Manager\
├── Launch.vbs                        <- Double-click to open the manager (no terminal window)
├── Launch.bat                        <- Use this for debugging (shows errors in console)
├── Windrose-Server-Manager.ps1       <- Main application
├── README.md
├── Backups\                          <- World backup zips (created on first backup)
├── player_history.txt                <- Running player join/leave log
├── R5\
│   ├── ServerDescription.json        <- Server config (managed by app)
│   └── Saved\
│       ├── Logs\R5.log               <- Live server log (read by Log tab)
│       └── SaveProfiles\             <- World save data (backed up by Tools tab)
├── Engine\                           <- UE5 engine files (copied from Steam)
└── WindroseServer.exe                <- Server launcher (copied from Steam)
```

---

## Troubleshooting

**Manager window closes immediately**
Run `Launch.bat` from a Command Prompt to see the error (use this instead of Launch.vbs when debugging). Common causes: PowerShell execution policy, missing .NET Framework.

**CPU shows 0% / RAM shows very low**
This resolves after the first 3-second watchdog cycle. The stats compare two snapshots apart to calculate CPU usage.

**Invite code shows "loading..."**
The server takes ~30-60 seconds to register with Windrose's backend after launch. The code auto-populates once ready and can also be found in `R5\ServerDescription.json`.

**Console commands not working**
The server must be started from the manager (not from `StartServerForeground.bat`) for stdin to be connected. If you started the server externally, restart it via the manager.

**Players can't connect from outside your network**
Confirm UDP 7777 and 7778 are forwarded on your router and that Windows Firewall allows `WindroseServer-Win64-Shipping.exe` through.

---

## Contributing

Issues and pull requests welcome. Please open an issue before starting significant work.

---

## License

MIT — see [LICENSE](LICENSE) for details.
