# IOIO Test Server

This folder contains a minimal local server for the VB6 simulator's
`/ioio/status` polling path.

The current VB6 code accepts either XML or JSON. This server returns JSON in the
simplest supported shape:

```json
{
  "analog_input": [123, 45, 67]
}
```

## Run

```bash
cd /Users/polinacoulter/SteamSimulator/ioio_server
python3 server.py
```

Default bind address:

- host: `127.0.0.1`
- port: `8080`

Endpoint:

- `GET /ioio/status`

## Test

```bash
curl http://127.0.0.1:8080/ioio/status
```

## Optional sample file

If `sample_data.json` exists, the server will load `analog_input` values from it
on each request. That makes it easy to edit test values without changing code.

Example:

```json
{
  "analog_input": [10, 20, 30, 40]
}
```

## VB6 integration

For local development, point the simulator at:

```vb
Public Const IOIO_STATUS_URL As String = "http://127.0.0.1:8080/ioio/status"
```

Current location:

- `SteamSim/steamsim-master/steamsim-master/STEAM_SIM_V31/VB Files/IOTypes.bas`
