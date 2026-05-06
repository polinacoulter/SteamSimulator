# Steam Simulator: Installation & Test Plan

## What you're installing

Three pieces that work together on the simulator's Windows XP machine:

1. **Python server** - middleman between the simulator and the hardware.
2. **Simulator** (`Steam_SimV32_00_00_Disabled_Profibus.exe`) - the steam
   simulator UI and model, with Profibus reads/writes disabled internally.
3. **I/O App** (`IOApp.exe`) - bridges the Python server and the Profibus
   cards.

The Raspberry Pi(s) on the network provide additional analog and digital
inputs that flow into the Python server.

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

## Step 2: Copy the project to XP

1. On a computer with the project, copy the `SteamSim` folder to your USB
   stick. It should contain:
   ```
   SteamSim/
       Steam_SimV32_00_00_Disabled_Profibus.exe   (the simulator)
       IOApp.exe                                  (the I/O bridge)
       IOApp.cfg                                  (Profibus block configuration)
       ioio_server/
           server.py                              (the Python middleman)
           server.cfg.json                        (upstream device config)
   ```
2. Plug the USB stick into the XP machine.
3. Copy the `SteamSim` folder to `C:\Source\` so the full path is
   `C:\Source\SteamSim\`.

The simulator's existing runtime data (`C:\Steam_Sim\Text\`,
`C:\Steam_Sim\Snapshots\`, `C:\Steam_Sim\Images\`, etc.) stays where it
is - several of those paths are hardcoded in the simulator's source and
the simulator expects them to exist. They're set up by the original
simulator installer and should not be moved or renamed.

### Optional: review configuration files

- `C:\Source\SteamSim\ioio_server\server.cfg.json` - list of upstream Pi
  devices the Python server polls. Default points at
  `192.168.100.202:8080`. Open in Notepad to add or edit.
- `C:\Source\SteamSim\IOApp.cfg` - Profibus block skip list. Defaults to
  skipping all analog input blocks (`cardA_ai_skip=all`,
  `cardB_ai_skip=all`) because the analog input Profibus hardware at CMA
  is currently broken. If a specific Profibus block needs to be excluded,
  add its address. Reread on every Start click in the I/O App.

---

## Step 3: Run the system

You'll have three windows open. Open them in this order.

### Window 1 - Python server

1. **Start -> Run -> `cmd`**.
2. Type: `cd C:\Source\SteamSim\ioio_server`
3. Type: `python server.py`
4. Leave this window open and visible. You should see output like:
   ```
   Starting Python HTTP server on http://127.0.0.1:8080
   ...
   Loaded 1 upstream device(s) from C:\Source\SteamSim\ioio_server\server.cfg.json
   Device polling [main_pi]: http://192.168.100.202:8080/ioio/status every 500 ms
   ```
5. Within ~5 seconds you should also see:
   ```
   Device polling [main_pi]: first response, N AIN + M DIN pins captured
   ```
   That line confirms the Python server has reached the Pi.

### Window 2 - Simulator

Double-click `C:\Source\SteamSim\Steam_SimV32_00_00_Disabled_Profibus.exe`.
The simulator's main window opens.

To start the model: **Run -> Cold** (or **Run -> Load Snapshot** to load a
saved state).

### Window 3 - I/O App

Double-click `C:\Source\SteamSim\IOApp.exe`. The form opens with **Status:
Stopped**.

Click **Start**. With Profibus cards present, you should see:
```
Card A initialized OK
Card B initialized OK
Starting I/O loop
```

If Profibus cards are not present (e.g., on a development laptop), you'll
get pop-up errors saying "IO Card A failed to initialise" and the same for
Card B. Dismiss both - this is expected without hardware. Status will show
"Running (no cards)" and the rest of the system continues to work.

---

## Step 4: Verify the integration

Five tests, in order. Each one verifies a specific link in the chain.

### Test 1: Pi -> Python server

In Firefox on the XP machine, navigate to:

```
http://192.168.100.202:8080/ioio/status
```

You should see a long JSON response listing analog and digital channels.

**If you see "Connection refused" or a blank page:** the Pi isn't reachable.
Verify the Pi is powered on and on the same closed LAN as the XP machine.

### Test 2: Python server is forwarding Pi data

In Firefox, navigate to:

```
http://127.0.0.1:8080/test/analog
```

You'll see a page of sliders. The right column (`ain`) should show
non-zero values for the Pi-controlled channels (typically 42-46).
Move a thruster or rudder on the simulator console - the corresponding
row should update within ~1 second.

**If the values are all zero:** the Python server isn't getting Pi data.
Check Window 1 - do you see the "first response, N AIN ..." line? If
not, the server hasn't reached the Pi yet.

### Test 3: Simulator is reading inputs

In the simulator: **Hardware -> HardWare Tests (DI/DO/AI/AO)**, then
click the **AI** tab. The list should show non-zero values for the Pi-
controlled channels.

Move a physical thruster on the simulator console. Within a couple of
seconds, the corresponding row in this list should update.

**If values are stuck at zero:** check Window 1 for `GET /ioio/status?t=...`
lines from the simulator. There should be one every ~300 ms.

### Test 4: Simulator is writing outputs

While the simulator's model is running, watch the I/O App log (Window 3).
Every couple of seconds you'll see a line like:

```
12:34:56  tick  ain[0]=0 aout[0]=128  din[0]=0 dout[0]=0
```

`aout[0]` should be non-zero and changing as the simulator runs.

**If aout[0] stays at zero forever:** check Window 1 for
`POST /ioio/outputs` lines. There should be one every ~300 ms while the
simulator is running.

### Test 5: Pi outputs are receiving server commands

Skip if Profibus cards aren't connected.

While the simulator is running, watch the physical lamps and gauges on
the simulator console. They should change as the simulator's model runs.
The Python server console should show occasional `POST /ioio/trigger`
lines when an output pin's value changes.

**If physical hardware doesn't respond:** confirm the I/O App status shows
"Running" (not "Running (no cards)") and that the log isn't repeating
"Card A AO write failed" or similar.

---

## Step 5: Shutting down

Stop in reverse order to avoid leaving processes running:

1. In the I/O App, click **Stop**, then close the window.
2. In the simulator, **File -> Exit**.
3. In the Python server console (Window 1), press **Ctrl+C**, then close
   the window.

---

## If something doesn't work

If you can't get one of the verification tests to pass, please collect the
following and send to support:

1. **Python server console output.** Click in Window 1, then
   **right-click -> Mark**, drag-select all text, press **Enter** to copy,
   then paste into an email.
2. **I/O App log screenshot.** Take a screenshot of Window 3 with the log
   visible.
3. **Simulator status bar screenshot.** Screenshot of the simulator's main
   window showing the status bar at the bottom.
4. **Direct Pi response.** In Firefox on XP, hit
   `http://192.168.100.202:8080/ioio/status` and copy the response into
   an email.
5. **Server response.** In Firefox on XP, hit
   `http://127.0.0.1:8080/ioio/status` and copy the response.
6. **Network diagnostics.** In `cmd` on XP:
   - Run `ipconfig` and copy the output.
   - Run `ping 192.168.100.202` and copy the output.

That's enough to figure out exactly where the chain is broken.
