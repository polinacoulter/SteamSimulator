# I/O App

Separate VB6 executable that bridges Profibus cards and the Python server.

## Architecture

```
   Full Simulator (Profibus disabled)
              | (HTTP)
              v
       Python Server  <->  Pi(s)
              ^
              | (HTTP)
              v
        I/O App  <->  Profibus cards
```

The simulator has Profibus disabled and uses the Python server for all I/O.
The I/O App is the inverse: Profibus enabled, talks to the same Python server
to push the inputs it reads from Profibus and pull the outputs it should write
to Profibus.

## Phase status

- [x] **Phase 1** — Skeleton form (Start/Stop buttons, log listbox, heartbeat
      timer). No HTTP, no Profibus.
- [x] **Phase 2** — HTTP plumbing. GET `/ioio/status` to pull aout/dout from
      server. POST `/ioio/inputs` to push ain/din to server. Modules:
      `IOTypes_App.bas`, `IOAppNet.bas`, `IOAppHandler.cls`.
- [x] **Phase 3** — Profibus reads/writes. `appio.bas` (DLL declarations) and
      `IOAppProfibus.bas` (wrapper functions copied from simulator's `IO.bas`,
      stripped of HTTP and orchestration). Form's Start tries `IO_Init` for
      both cards and MsgBoxes on failure. Timer body does per-card gated
      reads, HTTP push, HTTP pull, per-card gated writes.
- [x] **Phase 4** — Integration test: run both exes simultaneously, verify
      round-trip through the Python server. Plus a small fix: I/O App now
      skips `Send_Inputs_To_Python` when no Profibus cards are present, so
      it doesn't overwrite legitimate sources (Pi polling) with zeros.

## Files

| File | Purpose |
|---|---|
| `IOApp.vbp` | VB6 project definition |
| `IOAppForm.frm` | Main (only) form: Start/Stop, log, status, timer |
| `IOTypes_App.bas` | Arrays, address constants, card flags, URL constants |
| `IOAppNet.bas` | `Get_Outputs_From_Python` (async GET), `Send_Inputs_To_Python` (async POST) |
| `IOAppHandler.cls` | Async response handler — parses `aout=`/`dout=` into arrays |
| `IOAppProfibus.bas` | `Get_A_Input`/`Get_D_input`/`Set_A_Output`/`Set_D_Output` wrappers |
| `appio.bas` | DLL declarations (`IO_Init`, `IO_ReadIByte`, etc.) — verbatim from simulator |

## Opening on the XP box

1. Copy this `IOApp/` folder to the XP machine (anywhere on disk).
2. Launch VB6.
3. **File -> Open Project**, pick `IOApp.vbp`. All seven files load into the
   Project Explorer (Form, 1 class, 5 modules).
4. **Run -> Start** (or F5).
5. Click **Start**. Behavior depends on whether Profibus cards are present:
   - **No cards present** (current dev state): two MsgBoxes pop up
     ("IO Card A failed to initialise", then "IO Card B failed to initialise")
     — that matches what the original simulator did. Status flips to
     "Running (no cards)" and HTTP traffic still flows (zeros for ain/din).
   - **Cards present**: log shows "Card A initialized OK" / "Card B
     initialized OK" and the timer reads/writes Profibus alongside HTTP.
6. Click **Stop** to halt the timer.

## Build

**File -> Make IOApp.exe** produces the standalone executable.

## Verifying Phase 3 without Profibus hardware

The MsgBoxes are the proof. They mean `IO_Init` was called and returned
False — i.e., the Profibus code path is wired correctly and would work if
hardware were present. After dismissing the MsgBoxes, the timer keeps
ticking and HTTP traffic flows normally; you should see `GET /ioio/status`
and `POST /ioio/inputs` lines in the Python server console every ~500 ms.

## Phase 4 integration test

Goal: verify the simulator and I/O App can communicate through the Python
server.

### Setup

1. On the XP box: `cd ioio_server && python pages_server.py`. Leave it
   running, watching for log output.
2. Launch the simulator (Executive.exe) — it polls the server for inputs and
   pushes outputs.
3. Launch IOApp.exe. Click **Start**. Two MsgBoxes ("IO Card A failed",
   "IO Card B failed") — dismiss them. Status: "Running (no cards)".

You should now see four log streams:
- `pages_server.py` console: `GET /ioio/status` from simulator AND I/O App
  (twice as much traffic), plus `POST /ioio/outputs` from simulator.
- Simulator's behavior: same as before — pulls ain/din, pushes aout/dout.
- I/O App log: every ~2 s, `tick  ain[0]=N aout[0]=N  din[0]=N dout[0]=N`.
  The `aout[0]` and `dout[0]` should reflect what the simulator is currently
  outputting (e.g., gauge values, lamp states from the model).

### Test 1: simulator -> I/O App

In the simulator, do something that changes an output. Easy options:
- Run -> Cold (loads default state, model ticks, many outputs change).
- Move a control on an open form (e.g., a throttle handle) — that
  changes `A_OUTPUT[N]` for the corresponding channel.

Watch the I/O App log. The `aout[0]` value should match what the simulator
is outputting. Spot-check by calling out specific indices in the log line
(it currently shows index 0 only — extend the log line if you want more).

### Test 2: I/O App doesn't fight the Pi

With Pi polling enabled in `pages_server.py`, physical thruster/rudder
movements on the Pi should still propagate to the simulator's AI test form
even with the I/O App running. The Phase 4 fix (skip `Send_Inputs_To_Python`
when no cards) is what makes this work — without it, the I/O App would post
zeros every 500 ms and clobber the Pi's values.

To verify: with all three running (server, simulator, I/O App), wiggle a
thruster on the Pi. Simulator's AI test form should reflect the new value
within a couple of seconds. If you see values flickering between Pi values
and zeros, the fix didn't make it onto the XP box.

### Test 3: same Python server console handles both clients

After running for a minute, the server console should show:
- Steady stream of `GET /ioio/status?t=...&n=...` from both the simulator
  and the I/O App (cache-busted URLs from each).
- `POST /ioio/outputs` from the simulator (every ~300 ms).
- No `POST /ioio/inputs` from the I/O App (because no cards).
- Periodic `POST /api/set_analog` only when you drag browser sliders.

If all of that holds, Phase 4 passes and the architecture works end-to-end
modulo the Profibus side. Adding real cards later is the same code path,
just with non-zero Profibus reads/writes.
