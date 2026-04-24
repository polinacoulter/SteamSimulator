#!/usr/bin/env python3
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


HOST = "127.0.0.1"
PORT = 8080
BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DATA_PATH = BASE_DIR / "sample_data.json"


def load_payload():
    if SAMPLE_DATA_PATH.exists():
        with SAMPLE_DATA_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, dict) and isinstance(data.get("analog_input"), list):
            return data

    return {"analog_input": [0, 0, 0, 0]}


class IOIORequestHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path != "/ioio/status":
            self.send_error(404, "Not Found")
            return

        payload = load_payload()
        body = json.dumps(payload).encode("utf-8")

        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print("%s - - [%s] %s" % (self.address_string(), self.log_date_time_string(), format % args))


def main():
    server = ThreadingHTTPServer((HOST, PORT), IOIORequestHandler)
    print("Serving IOIO status on http://%s:%s/ioio/status" % (HOST, PORT))
    server.serve_forever()


if __name__ == "__main__":
    main()
