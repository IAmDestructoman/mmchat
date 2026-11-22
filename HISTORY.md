# MMChat history

This history was rebuilt in 2026 from a nested pile of backup folders: about 3,000 files, but only about 700 distinct. It also draws on version numbers in the `.vbp` files and compiled exes, PE build timestamps, file dates, and the Wayback Machine. Each version below is a tagged commit in this repo. Versions before 3.4.2 did not survive.

Development started around 1999, when the author was 14, on a 33.6k AOL dial-up with a copy of VB3/4. It ended in VB6.

## Server versions

### `v3.4.2` (August 2005)
- **Form:** a single, mostly headless form with a Microsoft Winsock control array.
- **Connection states:** `connect` → `need ver` → `need pw` → `auth`.
- **Licence check:** on startup the server sends `authnum*port*rand` to the auth server on port 4060. A `-1` reply opens the download page.
- **Settings:** the INI file is parsed by hand with `InStr`.
- **Working commands:** add, kick, login, channels, join, ccreate, info, gca, help, plog, setip, pmsg, who and log.
- **Stubs:** status, del, chpw, ca and ban. `gag` announces the gag but doesn't apply it.

### `v3.5.1` and `v3.5.1-dec` (November–December 2005)
- **Networking:** moves off the Winsock control to Oleg Gdalevich's pure-API `CSocket` class (version 1.0.12, published free on vbip.com in 2002), wrapped in the `Socket.ctl` user control.
- **Modified CSocket:** about a dozen `Err.Raise` calls in `SendData`, `PeekData`, `GetData`, `Listen` and `Connect` became `RaiseEvent OnError(...)`, so a socket failure reaches the server as an event instead of an unhandled error. `Option Explicit` was also switched off. A `wsstate()` helper and a `Debug.Print` were added in 2015.
- **Admin GUI:** chat name and port fields, online and banned lists, packet counters, gag settings, serve in/out options and minimise-to-tray.
- **Access control:** `banned.txt`; a server-wide "main password" step (new users wait in channel −1); a 15-second password timeout.
- **Flood control:** a real `gag()` with auto-gag.
- **Auth toggle:** the `auth` command switches login auth off. In `-dec` it only does so for about 2 minutes, then a timer switches it back on.
- **Side projects:** `Project1.vbp` is an attempt to run the server as an NT service (`NTSVC.ocx`). `MMChat2.vbp` is an empty side project.

### `v3.6.8` (April 2006)
This is the version in the `chathelp.jpg` screenshot.
- **Settings:** `ReadINI.bas` (Win32 profile API) replaces hand parsing. `loadsettings`/`savetoini`, and File › Setup with `@Bold`/`@Red`… colour tokens for the welcome text.
- **Version packet:** the client version now arrives as a `0x13…0xFF` packet at any time.
- **Bans:** cover usernames or IPs. Admins see banned users' connection attempts.
- **Channel access modes:** none, password, or permission level.
- **Channel logs:** capped at about 1,500 lines.
- **Commands:** chpw, ca, gca, ban, fjoin, remove and `help <command>` (from `Help/*.txt`) are implemented.
- **Gags:** keyed by IP, so reconnecting doesn't clear them.
- **Licence:** the check is skipped (`GoTo bypassit`). A demo mode is added.
- **Plugins:** the first loader (`Plugin.bas`) and the Points plugin. Nothing calls plugins yet.

### `v4.0.3` (July 2006)
- **Project:** renamed `ChatServer`. Globals move into `mmchat.bas`.
- **Plugin system:** COM DLLs in `Plugins\` load as `CreateObject("<dll>.Plugin")`, re-registered with `regsvr32` when that fails. They declare commands as `,cmd,level,`, return results routed by a leading `0`–`6` code, and get a re-entrancy lock.
- **New commands:** `plugins`, and `silent` for channels where only level 2+ may speak.
- **Accounts:** `.usr` account files become INI files. A File › Add User dialog is added.
- **Duplicate logins:** an authenticated login kicks the stale session.
- **Self-update:** `Form2` downloads a new exe, and a `.prg` rename swap installs it.
- **Licence and demo:** the licence check is back on. Demo builds randomly disconnect everyone at :00 and :30.

### `v4.1.2` (September 2006)
- **Command table:** `commands(1 To 24)` holds each command's name, level, parameter rule and help text. Checks are done once, centrally, and `[Perms]` in the INI can override levels.
- **Spam filter:** `replist.txt`, controlled by the `spamon` setting.
- **Plugins menu:** load all / unload all.
- **Public plugin SDK:** the next commit (Sept 26, 2006) adds the template published on the website in `example.zip`. Nobody else ever wrote a plugin.

### `v4.3.2` (2007)
The exe dates from September 2007 and the project file from December 2007. This is the version the "wrath" server was still running in 2010.
- **Removed:** the licence handshake and the main-password step.
- **Update check:** `checkupdate` reads `/mmchat/curversion.txt` from the website.
- **Plugins:** `init(App.Path)` is called on load.
- **Display:** an idle marker after 30 minutes, ANSI box tables for `who` and `channels`, and the connection type shown on join.

### `refactor-2008` branch (March 2008)
- **What changed:** all the form code was moved into `mmchat.bas` (`MMChatMOD`), `mmchat.frm` became an empty stub, and the sockets became a plain array of `CSocket`. This came alongside a VB6 → .NET upgrade assessment (reports dated 2008-03-30, not included).
- **Why it can't work:** VB6 arrays can't receive `WithEvents`, so no socket event ever fires and no connection would ever be handled. It also still references form controls. It never compiled.

### `v4.3.5` (January 2, 2013) and `v4.3.6` (July 24, 2013): binaries only
- **What survives:** only the compiled exes; their source was overwritten by later work. These commits change only `mmchat.exe`.
- **Deployment:** 4.3.6 ran as "town69" on port 4069.

### `v4.3.7-pre` (November 8, 2013): binary only
- **Origin:** recovered from the Wayback Machine (the website's `mmchat.exe`, captured April 2015). It was built seven minutes after `holofinder.dll`.
- **Has:** the plugin `Timer()` polling hook, added so holofinder could push announcements, the update prompt and silent `regsvr32`.
- **Missing:** the commands and service mode listed under `v4.3.7`.

### `v4.3.7` (January 26, 2015)
- **New commands:** `motd`, `setmotd` and `listusers`, bringing the table to `commands(1 To 27)`.
- **IP auth:** `ipauth` stores a list of `connectIP/reportedIP` pairs, so one account works from several places.
- **Service mode:** a `service` command-line switch runs it without the UI.
- **Update host:** hard-coded.
- **Builds today:** this source still compiles unmodified in VB6 (verified 2026).
- **May 2015:** a follow-up commit adds debugging edits to the socket class (`Debug.Print`, a `wsstate()` helper). They were never built.

### February 3, 2022
Fired up once more. The MOTD reads: *"Resurrection - 100s of hours of the worst coding by a teenager - but today that code lives again in 2022."*

## Plugins

| Plugin | Years | What it did |
|---|---|---|
| Points | Apr 2006 | Points and zones tracker backed by an Access/DAO database. Uses the pre-4.0 `main(text)` interface. |
| DarkNomads | May 2006 | Riddle answers, catacomb combinations and nearest clan towns from [darknomads.com](https://web.archive.org/web/2006/http://www.darknomads.com/) (see below). |
| message board | May 2006 | *"An almost exact replica of the Med Board"*: read, post, take. |
| telnet | Jun 2006 | Listens on the chat port + 1 and acts as a MudMaster client for raw-telnet users (`.setn name`, `.chat`). |
| Aim Notifier | Jul 2006 → Feb 2007 | Users register AIM screen names, and admins broadcast IMs. It logs into AIM over AOL's **TOC2** protocol, including the password-roasting and sign-on hash algorithms. |
| Battleship | Jul 2006 → Sep 2006 | Full two-player Battleship over chat, *"inspired by OMA's Battleship."* |
| connect4 | Jul 2006 | Connect Four, unfinished: the board is declared `As ibyte` (for `Byte`), so it never compiled and never shipped. |
| fbfly | Sep 2013 | "Faeball Fly": converts a bearing and distance into east/west and north/south moves. |
| holofinder | Nov 2013 → Aug 2014 | See below. |

### DarkNomads
darknomads.com was a Medievia help site with autoquest and zone walkthroughs, riddle answers, catacomb locations, nearest clan towns and player-submitted logs. It also offered a C++ MudMaster 2000 client DLL (not part of this project) with the same lookups.

The DarkNomads plugin brought those lookups into MMChat. It called the site's raw search endpoints (`/riddles/searchraw?name=…`, `/combs/searchraw?zone=…&room=…`, `/clantowns/searchraw?name=…`) and sent the same `User-Agent: Microsoft-ATL-Native/7` header as that C++ client (typed with stray spaces). The site's answers were already ANSI-formatted, so the plugin relays them almost unchanged.

### holofinder
Medievia published a live world-map GIF (`www.medievia.com/images/livemaps/livemapinfo.gif`). holofinder downloads it every 10 minutes (a 60-second timer, acting on every 10th tick) over a raw socket and:

1. **Decodes the GIF by hand**, written from CompuServe's GIF89a specification (a plain-text, RFC-style document) with no image library. That covers the header, logical screen descriptor, global colour table and image descriptor, then a full **variable-width LZW decoder**. The decoder handles clear codes, end-of-information, code-width growth up to 12 bits and the "KwKwK" case: a code the encoder created one step earlier, which the decoder hasn't added yet. That case comes up constantly in the map's long runs of solid colour.
2. **Scans the 2000×2000 pixel grid** for "holo" markers: a pixel of colour index 4 with exactly three orthogonal neighbours of the same colour and no diagonal ones.
3. **Measures the Manhattan distance** to the nearest road pixel (magenta) within 5 pixels.
4. **Diffs against the previous scan** and broadcasts new holos to the chat in ANSI colour, as age in minutes plus fly directions from the map centre ("NEW(E54 N45) Road:3"). Entries expire after 4 hours.

It was slow, because it handles bits as strings of `"0"`/`"1"`, rebuilds the LZW table with `ReDim Preserve` and calls `DoEvents` per code. But it worked. `Plugins/holofinder/gif/` is the earlier standalone version, which logged into the chat server as client "Holofind". Its original build didn't survive; the `gif.exe` there was compiled from the 2013 source with VB6 in 2026, unchanged. Turning it into a plugin is why the server gained the `Timer()` hook.

## The auth server (`authserver/`)

This was a licence server for MMChat installs, not a user-login server. Registered servers answered a numeric challenge on port 4060, and out-of-date builds were sent the update URL. A 2006 revision added a `DEMO` pass-through. The 2007 rewrite (port 4062, DNS verification through `DNS.CTL`) was never finished: it never sends a success reply and its `.vbp` is empty. The server stopped using licensing in 4.3.2.

## Out in the world

- **2006–2015:** downloads, plugins and the plugin SDK were published on the MMChat website ([archive-web/](archive-web/README.md)).
- **April 2010**, [Gammon forum thread 10198](https://www.gammon.com.au/forum/threads/10198.html), "Chat plugin in Medievia" (≈74k views):
  - **Problem:** a MUSHclient user couldn't log in to the "wrath" MMChat 4.3.2 server, though MudMaster could.
  - **Cause:** Nick Gammon used Wireshark to show that MUSHclient wraps chat text in ANSI colour codes. `pwauth` compares the raw text, so the password never matches.
  - **Workaround:** his alias sends the password without colour codes.
- **March 2025**, [Gammon forum thread 15230](https://www.gammon.com.au/forum/threads/15230.html), "wow, old times and things i never knew": the author finds the 2010 thread 15 years later and pins the bug on the one missing `stripansi` call.
