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

- [x] **Phase 1** — Skeleton form (Start/Stop buttons, log textbox, heartbeat
      timer). No HTTP, no Profibus yet. Goal: confirm the project compiles and
      the form opens.
- [ ] **Phase 2** — HTTP plumbing. GET `/ioio/status` to pull aout/dout from
      server. POST `/ioio/inputs` to push ain/din to server.
- [ ] **Phase 3** — Profibus reads/writes. Copy `Applicom.bas` and the relevant
      pieces of `IO.bas` from the simulator. Wire `IO_Init`, `Get_*`, `Set_*`
      into the timer loop.
- [ ] **Phase 4** — Run both exes simultaneously and verify the round-trip.

## Opening on the XP box

1. Copy this `IOApp/` folder to the XP machine (anywhere on disk).
2. Launch VB6.
3. **File -> Open Project**, navigate to the folder, select `IOApp.vbp`.
4. **Run -> Start** (or F5). The form should open with Status: Stopped.
5. Click **Start** -> log shows "tick" lines, status flips to Running.
6. Click **Stop** -> ticks stop, status returns to Stopped.

If that works, Phase 1 is done.

## Build

**File -> Make IOApp.exe** produces the standalone executable. Phase 4 will
involve running both `IOApp.exe` and the existing simulator `Executive.exe`
side by side.
