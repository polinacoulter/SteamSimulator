# Steam Simulator: Installation & Test Plan

The installation ships in two phases.

- **Phase 1 (now):** Python server + I/O App + Profibus hardware. You can
  install and verify these today using a browser to manually drive
  outputs and read inputs. No Pi, no simulator.
- **Phase 2 (later):** Raspberry Pi feed + simulator. Polina and Richard
  will ship updated `server.py`, the simulator EXE, and a refreshed
  `server.cfg.json` once the Pi-side software is updated to match the
  XML configuration in use at your site. Phase 1 components stay in
  place; Phase 2 layers on top.

This document covers Phase 1 in full. Phase 2 instructions are stubbed
in at the end and will be filled in when that software ships.

---

# PHASE 1: Python server + I/O App + Profibus

## What you're installing

1. **Python server** - middleman that holds I/O state in memory and
   exposes browser pages for manual input/output control.
2. **I/O App** (`IOApp.exe`) - bridges the Python server and the
   Profibus cards. Reads Profibus inputs and posts them to the server;
   reads outputs from the server and writes them to Profibus.

The simulator and the Raspberry Pi are **not** part of Phase 1.

---

## Step 1: Install Python on XP

Windows XP can only run **Python 3.4** (the last version that supports XP).

1. On a computer with internet, download:
   `https://www.python.org/ftp/python/3.4.4/python-3.4.4.msi`
2. Copy the installer to a USB stick.
3. Plug the USB stick into the XP machine.
4. Run the installer. **Important:** at the "Customize" screen, scroll to
   the bottom and enable **"Add python.exe to Path"**.
5. Click through to finish.

Verify: in `cmd`, run `python --version`. You should see `Python 3.4.4`.

---

## Step 2: Copy the Phase 1 files to XP

Copy these onto your USB stick:

```
IOApp.exe                                  (the I/O bridge)
IOApp.cfg                                  (Profibus block configuration)
io_server\
    server.py                              (the Python middleman)
    server.cfg.json                        (upstream device config)
```

On the XP machine, place them inside `C:\Steam_Sim\` so the final paths are:

```
C:\Steam_Sim\IOApp.exe
C:\Steam_Sim\IOApp.cfg
C:\Steam_Sim\io_server\server.py
C:\Steam_Sim\io_server\server.cfg.json
```

The simulator's existing runtime data (`C:\Steam_Sim\Text\`,
`C:\Steam_Sim\Snapshots\`, etc.) stays where it is; it'll be used in
Phase 2.

### Required: empty the upstream devices list

Open `C:\Steam_Sim\io_server\server.cfg.json` in Notepad and replace the
`devices` array with `[]`:

```json
{
  "devices": []
}
```

This stops the server from trying to poll the Pi during Phase 1. (The
default config points at a dev-side Pi that isn't reachable from your
network, and even if it were, it speaks a JSON format that the current
server expects rather than the XML your Pi emits. Phase 2 ships an
updated server that handles your Pi's format, along with the correct
device entry.)

### Optional: review IOApp.cfg

`C:\Steam_Sim\IOApp.cfg` is the Profibus block skip list. It defaults to
skipping all analog input blocks (`cardA_ai_skip=all`,
`cardB_ai_skip=all`) because the analog input Profibus hardware at CMA
is currently broken. If a specific Profibus block needs to be excluded,
add its address. Reread on every Start click in the I/O App.

---

## Step 3: Run the Phase 1 system

Two windows.

### Window 1 - Python server

1. **Start -> Run -> `cmd`**.
2. Type: `cd C:\Steam_Sim\io_server`
3. Type: `python server.py`
4. Leave this window open and visible. You should see output like:
   ```
   Starting Python HTTP server on http://127.0.0.1:8080
   ...
   Loaded 0 upstream device(s) from C:\Steam_Sim\io_server\server.cfg.json
   ```
   "0 devices" is correct for Phase 1 - we cleared the list in Step 2.

### Window 2 - I/O App

Double-click `C:\Steam_Sim\IOApp.exe`. The form opens with **Status:
Stopped**.

Click **Start**. With Profibus cards present, you should see:
```
Card A initialized OK
Card B initialized OK
Starting I/O loop
```

If Profibus cards are not present (e.g., on a development laptop), you'll
get pop-up errors saying "IO Card A failed to initialise" and the same
for Card B. Dismiss both - this is expected without hardware. Status
will show "Running (no cards)" and the rest of the system continues to
work, but Tests 1 and 3 below will not pass (they require real Profibus
hardware).

---

## Step 4: Verify Phase 1

Three tests, in order. They exercise the full Profibus link end-to-end:
hardware -> I/O App -> Python server -> browser, and back the other way.

### Test 1: I/O App is reading Profibus inputs

Watch the I/O App log (Window 2). Every couple of seconds you should
see a `tick` line:

```
12:34:56  tick  ain[0]=N aout[0]=N  din[0]=N dout[0]=N
```

The `ain` and `din` values should reflect whatever the Profibus cards
are currently reading from physical sensors and switches.

Wiggle a physical input (toggle a switch, move a sensor). Within a
couple of seconds the corresponding `ain[N]` or `din[N]` value in the
log should change.

**If values are stuck at zero:** confirm the I/O App status is
"Running" (not "Running (no cards)") and that the log isn't repeating
"Card A read failed" or similar.

### Test 2: Python server is receiving the Profibus inputs

In Firefox on the XP machine, navigate to:

```
http://127.0.0.1:8080/test/analog
```

The left column should show non-zero values for the Profibus-connected
channels. Move a physical analog input on the simulator console - the
corresponding row should update within ~1 second.

Then navigate to:

```
http://127.0.0.1:8080/test/digital
```

Toggle a physical switch - the corresponding row should flip from 0 to
1 (or back) within a second.

**If the values stay at zero:** check Window 1 for `POST /ioio/inputs`
or `POST /ioio/outputs` lines from the I/O App. If you don't see any,
the I/O App isn't reaching the server.

### Test 3: Profibus outputs respond to server-driven commands

Skip if Profibus cards aren't connected.

In Firefox, navigate to:

```
http://127.0.0.1:8080/test/digital
```

Find a `dout` row that maps to a known lamp or relay. Set its value to
1 and submit. Within ~1 second:
- The I/O App log should show a `tick` line with the corresponding
  `dout[N]=1`.
- The physical lamp/relay should turn on.

Set it back to 0 and confirm it turns off.

Repeat with an `aout` row on `http://127.0.0.1:8080/test/analog` to
exercise an analog output (gauge, motor, etc.). Try a few different
values - the physical device should move proportionally.

**If physical hardware doesn't respond:** confirm the I/O App status
shows "Running" (not "Running (no cards)") and that the log isn't
repeating "Card A AO write failed" or similar.

---

## Step 5: Shutting Phase 1 down

1. In the I/O App, click **Stop**, then close the window.
2. In the Python server console (Window 1), press **Ctrl+C**, then
   close the window.

---

## Phase 1 troubleshooting

If a verification test doesn't pass, please collect the following and
send to support:

1. **Python server console output.** Click in Window 1, then
   **right-click -> Mark**, drag-select all text, press **Enter** to
   copy, then paste into an email.
2. **I/O App log screenshot.** Take a screenshot of Window 2 with the
   log visible.
3. **Server response.** In Firefox on XP, hit
   `http://127.0.0.1:8080/ioio/status` and copy the response into an
   email.
4. **Network diagnostics.** In `cmd` on XP, run `ipconfig` and copy
   the output.

---

# PHASE 2: Raspberry Pi feed + simulator

Phase 2 adds two more components on top of Phase 1:

- **Raspberry Pi feed** - your Pi's XML-formatted status response gets
  parsed by the updated Python server and used to drive analog/digital
  inputs that the Profibus side doesn't cover.
- **Simulator** (`Steam_SimV32_00_00_Disabled_Profibus.exe`) - the steam
  simulator UI and model, with Profibus reads/writes disabled
  internally so they go through the Python server instead.

The Phase 1 components stay where they are; Phase 2 layers on top
without re-installing anything.

## What you're installing (Phase 2)

Three files replace or add to the Phase 1 install:

```
io_server\server.py                          (updated - parses Pi XML)
io_server\server.cfg.json                    (updated - points at your Pi)
Steam_SimV32_00_00_Disabled_Profibus.exe     (new - the simulator itself)
```

`IOApp.exe` and `IOApp.cfg` from Phase 1 stay exactly as installed.

---

## Step 1: Network - confirm the Pi is reachable

Before touching any files, from `cmd` on the XP machine:

```
ping <pi-ip>
```

You should get replies. Then:

```
curl http://<pi-ip>:8080/ioio/status
```

(If `curl` isn't installed on XP, point Firefox at the same URL.) You
should see an XML document listing pins. Copy the first ~20 lines of
that response into an email to support if anything looks off - it's
the exact input the updated `server.py` has to parse.

**If the ping fails:** the simulator XP host can't see the Pi on the
network. Fix that before going further; nothing else in Phase 2 will
work until the Pi is reachable.

---

## Step 2: Drop in the Phase 2 files

With the Python server and I/O App both **stopped** (see Phase 1
Step 5 if either is still running), copy from your USB stick:

- Overwrite `C:\Steam_Sim\io_server\server.py` with the Phase 2 version.
- Overwrite `C:\Steam_Sim\io_server\server.cfg.json` with the Phase 2
  version.
- Copy `Steam_SimV32_00_00_Disabled_Profibus.exe` into `C:\Steam_Sim\`.

Open `C:\Steam_Sim\io_server\server.cfg.json` in Notepad. Find:

```json
"url": "http://REPLACE_WITH_PI_IP:8080/ioio/status",
```

Replace `REPLACE_WITH_PI_IP` with the Pi's actual IP address (the one
that responded to the `ping` in Step 1). Save and close.

The simulator's existing runtime data (`C:\Steam_Sim\Text\`,
`C:\Steam_Sim\Snapshots\`, etc.) is used by the simulator EXE - it
must already be in place on this machine.

---

## Step 3: Launch order

Order matters. Bring components up one at a time so a failure at any
step is easy to isolate.

### Window 1 - Python server

1. **Start -> Run -> `cmd`**.
2. `cd C:\Steam_Sim\io_server`
3. `python server.py`

Expected output includes:

```
Starting Python HTTP server on http://127.0.0.1:8080
...
Loaded 1 upstream device(s) from C:\Steam_Sim\io_server\server.cfg.json
Device polling [cma_pi]: http://<pi-ip>:8080/ioio/status every 500 ms
Device polling [cma_pi]: first response, N AIN + M DIN pins captured
```

The "first response" line confirms the Pi is being reached and its
XML is being parsed. If you never see it, jump to Phase 2
troubleshooting below.

### Window 2 - I/O App

Double-click `C:\Steam_Sim\IOApp.exe`. Click **Start**. Same expected
output as Phase 1 - "Card A initialized OK", "Card B initialized OK",
"Starting I/O loop", and a stream of `tick` lines.

### Window 3 - Simulator

Double-click `C:\Steam_Sim\Steam_SimV32_00_00_Disabled_Profibus.exe`.
The simulator UI comes up. Give it a few seconds to settle and start
polling the server.

---

## Step 4: Verify Phase 2

Four tests, in order. Each one adds one more hop of the full loop.

### Test 4: Pi -> server

In Firefox on XP, `http://127.0.0.1:8080/test/analog`. The channels
your Pi covers should show non-zero values that update as physical
inputs on the Pi side change. Same drill on
`http://127.0.0.1:8080/test/digital` for digital inputs.

This is Phase 1's Test 2 but for the Pi-driven channels instead of
the Profibus-driven ones.

**If values stay at zero:** check Window 1 for the "first response"
line from Step 3. If it never appeared, the server isn't parsing the
Pi's XML successfully. Copy the raw response body (Step 1's `curl`
command) into an email to support.

### Test 5: Pi -> server -> simulator

Move a physical input on the Pi side. The corresponding
gauge/lamp/indicator in the simulator UI should update within ~1
second.

**If the browser page updates (Test 4 passed) but the simulator UI
doesn't:** the simulator isn't polling the server. Check that the
simulator's `IOIO_STATUS_URL` is pointed at `http://127.0.0.1:8080`
and not the old direct-Profibus path.

### Test 6: Simulator -> server -> I/O App -> Profibus

Drive an output from the simulator UI (start a pump, energize a
solenoid, whatever the training scenario calls for) whose target
device is on the Profibus side. Within ~1 second:

- Window 1 should show a `POST /ioio/outputs` line from the simulator.
- Window 2 should show a `tick` with the corresponding `dout[N]` or
  `aout[N]` value.
- The physical lamp/relay/gauge should respond.

This is the full write path end-to-end.

### Test 7: Snapshot load pushes expected inputs

Load a snapshot in the simulator. On load, the simulator pushes the
snapshot's expected `ain`/`din` values to the server so the
SnapCompare page starts clean.

Immediately after the load, open `http://127.0.0.1:8080/test/analog`
in Firefox. The values shown should match the snapshot's expected
inputs (until the Pi's next poll overwrites the Pi-covered channels,
which is fine - that's the normal running state).

---

## Step 5: Shutting Phase 2 down

Reverse the launch order:

1. Close the simulator window.
2. In the I/O App, click **Stop**, then close the window.
3. In the Python server console (Window 1), press **Ctrl+C**, then
   close the window.

---

## Phase 2 troubleshooting

Same collection procedure as Phase 1, plus:

5. **Raw Pi response.** From `cmd` on XP:
   `curl http://<pi-ip>:8080/ioio/status > pi_response.xml`
   Attach `pi_response.xml` to the email. This is what the server is
   trying to parse; a few characters out of place in the XML will
   break the parser silently.
6. **Server device log lines.** In Window 1, look for any line
   starting with `Device polling [cma_pi]:` and include all of them.

If Test 4 fails but Test 1/2/3 (from Phase 1) still pass, the
Profibus side is fine and the problem is Pi-side only - the two
paths are independent.
