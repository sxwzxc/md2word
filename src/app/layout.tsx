import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Markdown → Word 转换器",
  description: "在线将 Markdown 文档转换为格式精美的 Word (.docx) 文件，支持标题、加粗、代码、列表、表格等。",
  icons: {
    icon: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body className="antialiased">{children}</body>
    </html>
  );
}
