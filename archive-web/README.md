# archive-web: the original MMChat website, from the Wayback Machine

These files were recovered in November 2026 from the Internet Archive's Wayback Machine. They are what the MMChat download site served between 2006 and 2015. Each file was downloaded raw. The site's hostname and company name have been removed: in text they're replaced, in binaries they're overwritten at the same length with `mmchat.example.org` / `Dirocus`. So have the HTML author metadata and IDE workspace files.

| File | Captured | Notes |
|---|---|---|
| `website-2006/index.html` | 2006-08-22 | The first download page: *"put together in less then 5 mins in notepad, so deal with it."* |
| `website/index.html` | 2007-02-10 | Download page (Word-generated HTML, authored 2006-09-27). Unchanged in the 2007-05, 2008-11 and 2009-01 captures. |
| `website/plugins.html` | 2007-02-17 | Plugin list with descriptions. |
| `website/plugins.2009-01.html` | 2009-01-23 | The same page with `\` in place of `/` in its links. |
| `website/images/chatwho.jpg` | 2007-02-17 | `who` on a busy server (≈40 users across channels Main, stone and cpk), seen from the **MudMaster console client** (`mm` in `cmd.exe`). That client sat at 100% CPU busy-waiting on keyboard input, which hurt on the single-core laptops of the day. |
| `website/images/chathelp.jpg` | 2007-02-17 | `help` output from v3.6.8 in the same client, with Medievia auction spam scrolling underneath. |
| `website/mmchat.exe` | 2015-04-30 | Stamped **4.3.7**, but an earlier build (PE timestamp 2013-11-08) than the final 4.3.7. It has the update prompt, silent `regsvr32` and the plugin `Timer()` hook, but not `motd`, `listusers` or service mode. It exists nowhere else, and is the `v4.3.7-pre` tag. |
| `website/plugins/{aim,battleship,darknomads,message}.dll` | 2007-02-19 | The same builds as the DLLs in `Plugins/` as of `v4.3.2`. |
| `website/plugins/telnet.dll` | 2007-02-19 | A build that survives only here: built 2006-09-18 00:22, two and a half hours before the `Telnet.dll` in the repo. |
| `website/plugins/example.zip` | 2007-02-25 | The public plugin SDK, dated 2006-09-26. It is extracted into `website/plugins/example/` without the IDE workspace files. |

## What's in `example.zip`

`README.TXT`, the full sources of the message, battleship and telnet plugins (identical to `Plugins/` in this repo), and **`template/`**. The template is a commented skeleton plugin that explains the `GetName` / `GetCommands` / `main` / `getvar` contract and the `0`–`6` result codes. It survived only here, and is now also in `Plugins/template/`.

## Linked from the pages but never archived

`help.zip`, `images/main.jpg` (the third screenshot) and `curversion.txt`, the file the server's update check reads.

## The download pages' own words (2006/2007)

> Heres the deal, i wrote a program that supports zmud and mudmast clients or anyother client that can use one of those protocols like mush client for example. Anyway, its got lots of kick ass features. if you ask around, youll find someone that has used/seen it. so yeah if you wanna try it...
>
> REQUIREMENTS: 500mhz or better processor · 128+ megs ram · WINDOWS 2000,XP,2003 · STATIC OR SEMISTATIC IP ADDRESS
