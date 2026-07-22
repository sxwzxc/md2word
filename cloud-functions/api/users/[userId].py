"""
Python Cloud Function - Single Dynamic Parameter
api/users/[userId].py → GET /api/users/:userId
[userId].py captures the user identifier from the URL.
"""
import json
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):
    """Dynamic route handler — [userId] captures a single dynamic segment."""

    def do_GET(self):
        # Extract userId from the URL path
        parts = self.path.split('/')
        user_id = ''
        if len(parts) >= 4:
            user_id = parts[3].split('?')[0]

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({
            "userId": user_id,
            "name": f"User {user_id}",
            "email": f"{user_id}@example.com",
            "route": "/api/users/[userId]",
            "type": "single dynamic param",
        }).encode('utf-8'))
