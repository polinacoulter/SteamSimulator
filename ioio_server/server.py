#!/usr/bin/env python
import json
import os

try:
    from http.server import BaseHTTPRequestHandler, HTTPServer
except ImportError:
    from BaseHTTPServer import BaseHTTPRequestHandler, HTTPServer


HOST = "127.0.0.1"
PORT = 8080
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DATA_PATH = os.path.join(BASE_DIR, "sample_data.json")


def load_payload():
    if os.path.exists(SAMPLE_DATA_PATH):
        with open(SAMPLE_DATA_PATH, "r") as handle:
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
        values = payload.get("analog_input", [])

        parts = ['<?xml version="1.0" encoding="UTF-8"?>', "<pins>"]
        for index, value in enumerate(values):
            parts.append('<pin name="ain%d" calibrated="%d"/>' % (index, int(value)))
        parts.append("</pins>")
        body = "".join(parts).encode("utf-8")

        self.send_response(200)
        self.send_header("Content-Type", "text/xml; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print("%s - - [%s] %s" % (self.address_string(), self.log_date_time_string(), format % args))


def main():
    server = HTTPServer((HOST, PORT), IOIORequestHandler)
    print("Serving IOIO status on http://%s:%s/ioio/status" % (HOST, PORT))
    server.serve_forever()


if __name__ == "__main__":
    main()
