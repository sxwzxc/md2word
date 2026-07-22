"""
Python Cloud Function - Catch-All Route
api/files/[[path]].py → GET /api/files/*path
[[path]].py is a catch-all route that matches any number of path segments.
e.g. /api/files/docs/guide/intro.md → path = "docs/guide/intro.md"
"""
import json
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):
    """Catch-all route handler — [[path]] catches all remaining segments."""

    def do_GET(self):
        # Extract the catch-all path after /api/files/
        file_path = ''
        prefix = '/api/files/'
        path = self.path.split('?')[0]
        if path.startswith(prefix):
            file_path = path[len(prefix):]

        # Determine file extension
        ext = ''
        if '.' in file_path:
            ext = file_path[file_path.rfind('.'):]

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({
            "path": file_path,
            "extension": ext,
            "segments": file_path.split('/') if file_path else [],
            "route": "/api/files/[[path]]",
            "type": "catch-all route",
        }).encode('utf-8'))
