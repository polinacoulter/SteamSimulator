# SteamSimulator

Integration work on the Cal Maritime steam plant simulator. Three components
talking through a Python HTTP middleman, replacing direct Profibus I/O so the
simulator can run on machines without the Applicom hardware.

## Architecture

```
   Full Simulator (Profibus disabled, VB6)
              |  HTTP
              v
       Python Server  <->  Raspberry Pi (real hardware)
              ^
              |  HTTP
              v
        I/O App (separate VB6 EXE)  <->  Profibus cards
```

The Python server (`ioio_server/server.py`) maintains in-memory I/O state.
The full simulator and the I/O App both talk to it; they never talk to
each other directly. The Pi pushes physical sensor readings into the
server's state on a configurable polling schedule.

## Repo layout

| Path | What lives there |
|---|---|
| `steamsim-master/` | The full simulator's VB6 source tree (canonical name carried over from the original distribution). Modified files: `IO.bas`, `IOTypes.bas`, `IOIOReadyStateHandler.cls`, `TakeSnapForm.frm`, `Host.Frm` under `STEAM_SIM_V31/VB Files/`. |
| `IOApp/` | The new I/O App VB6 project. Standalone executable that bridges Profibus and the Python server. Has its own README inside. |
| `ioio_server/` | The Python middleman: `server.py`, `server.cfg.json`, README. |

## What got built (milestone-by-milestone)

- **Read analog inputs from Python server** - Simulator reads AIN values from
  the server instead of Profibus. Same for digital inputs.
- **Write outputs to Python server** - Simulator pushes A_OUTPUT and D_OUTPUT
  to the server every poll cycle.
- **On snapshot load, push expected inputs** - When a snapshot loads, the
  simulator pushes the snapshot's expected ain/din state to the server so
  the SnapCompare page starts clean and surfaces real changes as mismatches.
- **Python server reads from Pi** - Background polling thread per device,
  parses the Pi's JSON, captures both AIN and DIN. Configurable via
  `ioio_server/server.cfg.json`.
- **Python server writes to Pi** - For each output pin the Pi exposes,
  POST `/ioio/trigger` when the value changes.
- **I/O App** - New separate VB6 EXE. Reads Profibus inputs and POSTs them
  to the server; reads outputs from the server and writes them to Profibus.
  Per-block skip configuration in `IOApp/IOApp.cfg` (a known-broken Profibus
  block can be excluded without touching the rest of the bus).

## Deferred / pending hardware

- **Synchroscope** - Deferred until further work on the Java code.
- **End-to-end Profibus verification** - I/O App code is wired but final
  proof requires real Profibus cards at CMA.

## Deployment

For installation and verification on the Cal Maritime simulator host, see
`TestPlan.md` (in this folder).

## Constraints worth knowing if you pick this up

- **No internet on XP, manual file transfer.** The simulator host is a
  Windows XP machine that must stay off the public internet (no security
  patches since 2014). Files move Mac to USB to XP.
- **VB6 wants CRLF.** Mac saves LF; convert before copying VB6 sources over.
- **VB6 chokes on non-ASCII.** Em-dashes specifically broke .frm parsing.
  Stay ASCII.
- **VB6 IDE hides `Attribute` lines** but they're required in the .frm /
  .cls / .bas files on disk. When patching files manually, preserve them.
- **Multiline VB.TextBox needs a companion .frx file.** Use VB.ListBox to
  dodge this.
- **MSXML2.XMLHTTP shares WinINet's GET cache.** Without a per-request
  cache-buster query string, identical GETs to `/ioio/status` are silently
  served from cache and never hit the network. Already fixed in
  `IO.bas/Request_IOIO_Data` but worth knowing if it crops up elsewhere.
