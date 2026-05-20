# IOIO Test Server

This folder contains a minimal local server for the VB6 simulator's
`/ioio/status` polling path.

The current VB6 code accepts either XML or JSON. This server returns XML in the
shape `IOIOReadyStateHandler.cls` parses natively via MSXML2:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<pins>
  <pin name="ain0" calibrated="123"/>
  <pin name="ain1" calibrated="45"/>
  <pin name="ain2" calibrated="67"/>
</pins>
```

XML is used on the wire because VB6 on XP has no built-in JSON parser;
MSXML2 handles XML natively. `sample_data.json` stays as the editable
source — the server reads integers from it and renders XML on each
request.

## Run

```bash
cd /Users/polinacoulter/SteamSimulator/io_server
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
