"""
Python Cloud Function - Nested Dynamic Parameters
api/users/[userId]/posts/[postId].py → GET /api/users/:userId/posts/:postId
Nested dynamic params — both [userId] and [postId] are captured.
"""
import json
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):
    """Nested dynamic route handler — captures [userId] and [postId]."""

    def do_GET(self):
        # Extract userId and postId from the URL path
        parts = self.path.split('/')
        user_id = ''
        post_id = ''
        if len(parts) >= 6:
            user_id = parts[3].split('?')[0]
            post_id = parts[5].split('?')[0]

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('X-Powered-By', 'Python Cloud Function')
        self.end_headers()
        self.wfile.write(json.dumps({
            "userId": user_id,
            "postId": post_id,
            "title": f"Post {post_id} by User {user_id}",
            "content": "This is a sample post content.",
            "route": "/api/users/[userId]/posts/[postId]",
            "type": "multiple dynamic params",
        }).encode('utf-8'))
