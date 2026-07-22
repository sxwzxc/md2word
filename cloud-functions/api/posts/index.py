"""
Python Cloud Function - Index Route
api/posts/index.py → GET /api/posts
index.py serves as the default handler for the /api/posts/ directory.
"""
import json
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):
    """Index route handler — index.py serves as the default for a directory."""

    def do_GET(self):
        posts = [
            {"id": 1, "title": "Getting Started with Python Functions", "slug": "getting-started"},
            {"id": 2, "title": "File-Based Routing in Python", "slug": "file-based-routing"},
            {"id": 3, "title": "Deploy Python to EdgeOne Pages", "slug": "deploy-python-edgeone"},
        ]

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({
            "posts": posts,
            "total": len(posts),
            "route": "/api/posts",
            "type": "index route (index.py)",
        }, ensure_ascii=False).encode('utf-8'))
