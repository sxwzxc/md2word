import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Markdown → Word 转换器",
  description: "上传 Markdown 文件，一键转换为格式优雅的 Word (.docx) 文档。",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <head>
        <link rel="icon" href="/favicon.svg" />
      </head>
      <body
        className="antialiased"
      >
        {children}
      </body>
    </html>
  );
}
