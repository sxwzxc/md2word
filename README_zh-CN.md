# Python Cloud Functions on EdgeOne Pages - Handler Mode

一个基于 Next.js + Tailwind CSS 的函数请求演示网站，展示如何在 EdgeOne Pages 上使用 Handler 模式部署 Python 云函数，支持基于文件的路由系统。

## 🚀 特性

- **文件路由系统**：Python 处理函数文件直接映射为 API 端点，类似 Next.js 的文件路由
- **现代化 UI 设计**：采用黑底白字主题，使用 #1c66e5 作为点缀色
- **实时 API 演示**：集成 Python 后端，支持所有路由类型的交互式调用测试
- **多种路由模式**：支持静态路由、索引路由、动态参数、嵌套动态参数和 Catch-all 路由
- **TypeScript 支持**：完整的类型定义和类型安全

## 🛠️ 技术栈

### 前端
- **Next.js 15** - React 全栈框架
- **React 19** - 用户界面库
- **TypeScript** - 类型安全的 JavaScript
- **Tailwind CSS 4** - 实用优先的 CSS 框架

### UI 组件
- **Lucide React** - 精美的图标库
- **class-variance-authority** - 组件样式变体管理
- **clsx & tailwind-merge** - CSS 类名合并工具

### 后端
- **Python 3.9+** - 云函数运行时
- **Handler 模式** - EdgeOne Pages 上基于文件的 Python 函数路由

## 📁 项目结构

```
python-handler-template/
├── cloud-functions/                # Python 云函数源码
│   ├── hello.py                   # 静态路由 → GET /hello
│   └── api/
│       ├── posts/
│       │   └── index.py           # 索引路由 → GET /api/posts
│       ├── users/
│       │   ├── [userId].py        # 动态参数 → GET /api/users/:userId
│       │   └── [userId]/
│       │       └── posts/
│       │           └── [postId].py # 嵌套动态参数 → GET /api/users/:userId/posts/:postId
│       └── files/
│           └── [[path]].py        # Catch-all → GET /api/files/*path
├── src/
│   ├── app/                       # Next.js App Router
│   │   ├── globals.css           # 全局样式
│   │   ├── layout.tsx            # 根布局
│   │   └── page.tsx              # 主页面（API 演示界面）
│   ├── components/               # React 组件
│   │   └── ui/                   # UI 基础组件
│   │       ├── button.tsx        # 按钮组件
│   │       └── card.tsx          # 卡片组件
│   └── lib/                      # 工具函数
│       └── utils.ts              # 通用工具
├── public/                        # 静态资源
├── package.json                   # 项目配置
└── README.md                     # 项目文档
```

## 🚀 快速开始

### 环境要求

- Node.js 18+
- npm 或 yarn
- Python 3.9+（本地开发需要）

### 安装依赖

```bash
npm install
# 或
yarn install
```

### 开发模式

```bash
edgeone pages dev
```

访问 [http://localhost:8088](http://localhost:8088) 查看应用。

### 构建生产版本

```bash
edgeone pages build
```

## 🎯 核心功能

### 1. 基于文件的 Python 路由

`cloud-functions/` 目录直接映射为 API 路由：

| 文件 | 路由 | 模式 |
|------|------|------|
| `hello.py` | `GET /hello` | 静态路由 |
| `api/posts/index.py` | `GET /api/posts` | 索引路由 |
| `api/users/[userId].py` | `GET /api/users/:userId` | 动态参数 |
| `api/users/[userId]/posts/[postId].py` | `GET /api/users/:userId/posts/:postId` | 嵌套动态参数 |
| `api/files/[[path]].py` | `GET /api/files/*path` | Catch-all 路由 |

### 2. 交互式 API 演示

- 点击 "Call" 按钮实时测试每个 API 端点
- 查看 JSON 响应结果
- 可展开查看源文件路径

### 3. Handler 约定

每个 Python 文件导出标准的 `handler` 类：

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

## 🔧 配置说明

### Tailwind CSS 配置
项目使用 Tailwind CSS 4，支持自定义颜色变量：

```css
:root {
  --primary: #1c66e5;        /* 主色调 */
  --background: #000000;     /* 背景色 */
  --foreground: #ffffff;     /* 前景色 */
}
```

### 组件样式
使用 `class-variance-authority` 管理组件样式变体，支持多种预设样式。

## 📚 文档入口

- **EdgeOne Pages 官方文档**：[https://pages.edgeone.ai/document/python](https://pages.edgeone.ai/document/python)
- **Next.js 文档**：[https://nextjs.org/docs](https://nextjs.org/docs)
- **Tailwind CSS 文档**：[https://tailwindcss.com/docs](https://tailwindcss.com/docs)
- **Python 文档**：[https://docs.python.org/3](https://docs.python.org/3)

## 🚀 部署指南

### EdgeOne Pages 部署

1. 将代码推送到 GitHub 仓库
2. 在 EdgeOne Pages 控制台创建新项目
3. 选择 GitHub 仓库作为源
4. 配置构建命令：`edgeone pages build`
5. 部署项目

### Python 云函数配置

在项目根目录创建 `cloud-functions/` 文件夹，添加 Python 处理函数：

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

## 部署

[![Deploy with EdgeOne Pages](https://cdnstatic.tencentcs.com/edgeone/pages/deploy.svg)](https://console.cloud.tencent.com/edgeone/pages/new?from=github&template=python-handler-template)

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](https://github.com/github/choosealicense.com/blob/gh-pages/_licenses/mit.txt) 文件了解详情。
