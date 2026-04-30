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
- [ ] **Phase 4** — Run both exes simultaneously and verify the round-trip.

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

## Phase 4 (next)

Run `Executive.exe` (the simulator with Profibus disabled) and `IOApp.exe`
side by side, both pointed at the same `pages_server.py`. Confirm the
round-trip: simulator writes `A_OUTPUT[0] = 200` -> arrives at I/O App's
`A_OUTPUT[0]` via HTTP -> in a real deployment, would land on Profibus.
Without real cards, just verify both ends see the same values.
