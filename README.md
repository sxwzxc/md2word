# Python Cloud Functions on EdgeOne Pages - Handler Mode

A function request demonstration website based on Next.js + Tailwind CSS, showcasing how to deploy Python Cloud Functions using Handler Mode on EdgeOne Pages with file-based routing.

## 🚀 Features

- **File-Based Routing**: Python handler files map directly to API endpoints, just like Next.js file routing
- **Modern UI Design**: Adopts black background with white text theme, using #1c66e5 as accent color
- **Real-time API Demo**: Integrated Python backend with interactive API call testing for all route types
- **Multiple Route Patterns**: Supports static, index, dynamic, nested dynamic, and catch-all routes
- **TypeScript Support**: Complete type definitions and type safety

## 🛠️ Tech Stack

### Frontend
- **Next.js 15** - React full-stack framework
- **React 19** - User interface library
- **TypeScript** - Type-safe JavaScript
- **Tailwind CSS 4** - Utility-first CSS framework

### UI Components
- **Lucide React** - Beautiful icon library
- **class-variance-authority** - Component style variant management
- **clsx & tailwind-merge** - CSS class name merging utilities

### Backend
- **Python 3.9+** - Cloud Functions runtime
- **Handler Mode** - File-based routing for Python functions on EdgeOne Pages

## 📁 Project Structure

```
python-handler-template/
├── cloud-functions/                # Python Cloud Functions source
│   ├── hello.py                   # Static route → GET /hello
│   └── api/
│       ├── posts/
│       │   └── index.py           # Index route → GET /api/posts
│       ├── users/
│       │   ├── [userId].py        # Dynamic param → GET /api/users/:userId
│       │   └── [userId]/
│       │       └── posts/
│       │           └── [postId].py # Nested params → GET /api/users/:userId/posts/:postId
│       └── files/
│           └── [[path]].py        # Catch-all → GET /api/files/*path
├── src/
│   ├── app/                       # Next.js App Router
│   │   ├── globals.css           # Global styles
│   │   ├── layout.tsx            # Root layout
│   │   └── page.tsx              # Main page (API demo UI)
│   ├── components/               # React components
│   │   └── ui/                   # UI base components
│   │       ├── button.tsx        # Button component
│   │       └── card.tsx          # Card component
│   └── lib/                      # Utility functions
│       └── utils.ts              # Common utilities
├── public/                        # Static assets
├── package.json                   # Project configuration
└── README.md                     # Project documentation
```

## 🚀 Quick Start

### Requirements

- Node.js 18+
- npm or yarn
- Python 3.9+ (for local development)

### Install Dependencies

```bash
npm install
# or
yarn install
```

### Development Mode

```bash
edgeone pages dev
```

Visit [http://localhost:8088](http://localhost:8088) to view the application.

### Build Production Version

```bash
edgeone pages build
```

## 🎯 Core Features

### 1. File-Based Python Routing

The `cloud-functions/` directory maps directly to API routes:

| File | Route | Pattern |
|------|-------|---------|
| `hello.py` | `GET /hello` | Static route |
| `api/posts/index.py` | `GET /api/posts` | Index route |
| `api/users/[userId].py` | `GET /api/users/:userId` | Dynamic parameter |
| `api/users/[userId]/posts/[postId].py` | `GET /api/users/:userId/posts/:postId` | Nested dynamic parameters |
| `api/files/[[path]].py` | `GET /api/files/*path` | Catch-all route |

### 2. Interactive API Demo

- Click "Call" to test each API endpoint in real-time
- View JSON response with syntax highlighting
- Expandable source file path display

### 3. Handler Convention

Each Python file exports a standard `handler` class:

```python
import json
import time
from http.server import BaseHTTPRequestHandler

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({
            "message": "Hello from Python Cloud Function!",
            "timestamp": time.time()
        }).encode('utf-8'))
```

## 🔧 Configuration

### Tailwind CSS Configuration
The project uses Tailwind CSS 4 with custom color variables:

```css
:root {
  --primary: #1c66e5;        /* Primary color */
  --background: #000000;     /* Background color */
  --foreground: #ffffff;     /* Foreground color */
}
```

### Component Styling
Uses `class-variance-authority` to manage component style variants with multiple preset styles.

## 📚 Documentation

- **EdgeOne Pages Official Docs**: [https://pages.edgeone.ai/document/python](https://pages.edgeone.ai/document/python)
- **Next.js Documentation**: [https://nextjs.org/docs](https://nextjs.org/docs)
- **Tailwind CSS Documentation**: [https://tailwindcss.com/docs](https://tailwindcss.com/docs)
- **Python Documentation**: [https://docs.python.org/3](https://docs.python.org/3)

## 🚀 Deployment Guide

### EdgeOne Pages Deployment

1. Push code to GitHub repository
2. Create new project in EdgeOne Pages console
3. Select GitHub repository as source
4. Configure build command: `edgeone pages build`
5. Deploy project

### Python Cloud Functions Configuration

Create `cloud-functions/` folder in project root and add Python handler files:

```python
# cloud-functions/hello.py
import json
import time
from http.server import BaseHTTPRequestHandler

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({
            "message": "Hello from Python!",
            "timestamp": time.time()
        }).encode('utf-8'))
```

## Deploy

[![Deploy with EdgeOne Pages](https://cdnstatic.tencentcs.com/edgeone/pages/deploy.svg)](https://edgeone.ai/pages/new?from=github&template=python-handler-template)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/github/choosealicense.com/blob/gh-pages/_licenses/mit.txt) file for details.
