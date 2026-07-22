"""
Python Cloud Function - Static Route
hello.py → GET /hello
"""
import json
import time
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):
    """Static route handler — file name maps directly to path."""

    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({
            "message": "Hello from Python Functions on EdgeOne Pages!",
            "route": "/hello",
            "type": "static route",
        }).encode('utf-8'))
