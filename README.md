# Markdown → Word 转换器

一个在线 Markdown 转 Word (.docx) 工具，基于 Next.js 前端 + Python 云函数后端，部署在 EdgeOne Pages 上。

上传 `.md` 文件，一键下载格式精美的 `.docx` 文档。

## 特性

- **纯 Python 后端** — 使用 python-docx 生成 Word 文档，无需 LibreOffice / Pandoc
- **完全干净的格式** — 不使用任何 heading / list 样式，避免 WPS / Word 中出现项目符号方框点
- **CJK 字体支持** — 中文正文使用微软雅黑，标题使用黑体，不会出现缺字方块
- **两种排版模式** — mode=1 基础模式，mode=2 精美模式（1.5 倍行距、配色标题、代码块边框、斑马纹表格、可点击超链接）
- **拖拽上传** — 支持拖拽 `.md` / `.markdown` / `.txt` 文件，最大 10 MB
- **支持完整 Markdown 语法** — 标题、粗体、斜体、删除线、行内代码、代码块、有序/无序列表、引用块、表格、超链接、分隔线

## 技术栈

| 层 | 技术 |
|---|---|
| 前端 | Next.js 15、React 19、TypeScript、Tailwind CSS 4 |
| 后端 | Python 3.9+、python-docx、EdgeOne Pages Handler Mode |
| 部署 | EdgeOne Pages（推送 GitHub `main` 分支自动部署） |

## 项目结构

```
md2word/
├── cloud-functions/               # Python 云函数
│   ├── api/convert/
│   │   └── md-to-docx.py         # POST /api/convert/md-to-docx
│   └── requirements.txt
├── src/                           # Next.js 前端
│   ├── app/
│   │   ├── layout.tsx            # 根布局 + favicon
│   │   ├── page.tsx              # 主页面（上传 / 转换 / 下载）
│   │   └── globals.css           # 全局样式
│   ├── components/ui/
│   │   ├── button.tsx
│   │   └── card.tsx
│   └── lib/utils.ts
├── public/
│   ├── favicon.svg               # 站点图标
│   └── eo-logo-blue.svg          # EdgeOne logo
└── package.json
```

## 快速开始

### 环境要求

- Node.js 18+
- Python 3.9+（本地开发云函数时需要）

### 安装与开发

```bash
npm install
edgeone pages dev
```

访问 http://localhost:8088

### API 用法

```bash
# 基础模式
curl -X POST 'https://md2word.sxwzxc.cn/api/convert/md-to-docx?mode=1' \
  --data-binary '# 标题

正文内容。'

# 精美模式
curl -X POST 'https://md2word.sxwzxc.cn/api/convert/md-to-docx?mode=2' \
  --data-binary '# 标题

正文内容。'
```

响应为 JSON，`data` 字段为 base64 编码的 `.docx` 文件：

```json
{
  "ok": true,
  "mode": 1,
  "ver": "v7-nokeep",
  "filename": "converted.docx",
  "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "size": 32882,
  "data": "UEsDBBQ..."
}
```

## Word 文档格式说明

转换后的 `.docx` 经过深度清理，确保在 Microsoft Word 和 WPS Office 中都能正确显示：

- **不使用 heading / title / list 样式** — 标题以普通段落 + 显式 run 格式实现，避免样式继承导致的格式问题
- **不使用 numbering 定义** — 列表使用手动文本前缀（`•` / `1.`），numbering.xml 为空
- **不使用 keepNext / keepLines** — 避免 WPS 中显示为段落前的方框标记
- **显式字体设置** — 所有 *Theme 字体引用已移除，theme1.xml 的空 eastAsia 字体槽已填充

## 部署

[![Deploy with EdgeOne Pages](https://cdnstatic.tencentcs.com/edgeone/pages/deploy.svg)](https://console.cloud.tencent.com/edgeone/pages/new?from=github)

1. Fork 本仓库
2. 在 EdgeOne Pages 控制台创建新项目，选择该 GitHub 仓库
3. 构建命令：`edgeone pages build`
4. 推送到 `main` 分支即自动部署

## 许可证

MIT License
