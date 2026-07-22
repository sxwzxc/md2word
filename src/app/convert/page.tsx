"use client"

import { useState, useRef, useCallback } from "react"
import Link from "next/link"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import {
  Upload,
  FileText,
  Download,
  ArrowLeft,
  Loader2,
  CheckCircle2,
  AlertCircle,
  FileUp,
} from "lucide-react"

const MAX_SIZE = 10 * 1024 * 1024 // 10 MB

export default function ConvertPage() {
  const [fileName, setFileName] = useState<string | null>(null)
  const [markdownContent, setMarkdownContent] = useState<string>("")
  const [status, setStatus] = useState<"idle" | "converting" | "done" | "error">("idle")
  const [errorMsg, setErrorMsg] = useState<string>("")
  const [docxUrl, setDocxUrl] = useState<string | null>(null)
  const [isDragging, setIsDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const objectUrlRef = useRef<string | null>(null)

  const handleFile = useCallback((file: File) => {
    // Revoke any previous object URL
    if (objectUrlRef.current) {
      URL.revokeObjectURL(objectUrlRef.current)
      objectUrlRef.current = null
    }

    setDocxUrl(null)
    setStatus("idle")
    setErrorMsg("")

    const name = file.name.toLowerCase()
    if (!name.endsWith(".md") && !name.endsWith(".markdown") && !name.endsWith(".txt")) {
      setErrorMsg("Please select a Markdown file (.md, .markdown, .txt)")
      setStatus("error")
      return
    }

    if (file.size > MAX_SIZE) {
      setErrorMsg("File too large. Maximum size is 10 MB.")
      setStatus("error")
      return
    }

    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      setMarkdownContent(String(e.target?.result ?? ""))
    }
    reader.onerror = () => {
      setErrorMsg("Failed to read file.")
      setStatus("error")
    }
    reader.readAsText(file)
  }, [])

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file) handleFile(file)
  }

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) handleFile(file)
  }

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(true)
  }

  const onDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
  }

  const handleConvert = async () => {
    if (!markdownContent.trim()) {
      setErrorMsg("Markdown content is empty.")
      setStatus("error")
      return
    }

    setStatus("converting")
    setErrorMsg("")

    try {
      const res = await fetch("/api/convert/md-to-docx", {
        method: "POST",
        headers: { "Content-Type": "text/plain; charset=utf-8" },
        body: markdownContent,
      })

      if (!res.ok) {
        let msg = `Server returned ${res.status}`
        try {
          const errJson = await res.json()
          msg = errJson.error || msg
        } catch {
          // response wasn't JSON
        }
        throw new Error(msg)
      }

      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      objectUrlRef.current = url
      setDocxUrl(url)
      setStatus("done")
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : "Conversion failed.")
      setStatus("error")
    }
  }

  const downloadName = fileName
    ? fileName.replace(/\.(md|markdown|txt)$/i, "") + ".docx"
    : "converted.docx"

  const handleReset = () => {
    if (objectUrlRef.current) {
      URL.revokeObjectURL(objectUrlRef.current)
      objectUrlRef.current = null
    }
    setFileName(null)
    setMarkdownContent("")
    setStatus("idle")
    setErrorMsg("")
    setDocxUrl(null)
    if (inputRef.current) inputRef.current.value = ""
  }

  return (
    <div className="min-h-screen bg-black text-white relative overflow-hidden">
      {/* Grid Background */}
      <div className="grid-background" />

      {/* Background Gradient Orbs */}
      <div className="gradient-orb gradient-orb-primary w-[600px] h-[600px] -top-[200px] -left-[150px] animate-pulse-glow" />
      <div className="gradient-orb gradient-orb-secondary w-[400px] h-[400px] top-[40%] -right-[100px] animate-pulse-glow animation-delay-200" />

      {/* Header */}
      <header className="header-border relative z-10">
        <div className="container mx-auto px-6 py-4">
          <div className="flex items-center justify-between">
            <Link
              href="/"
              className="flex items-center gap-2 text-gray-400 hover:text-[#3776AB] transition-colors"
            >
              <ArrowLeft className="w-4 h-4" />
              <span className="text-sm">Back to Home</span>
            </Link>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-6 py-16 relative z-10">
        <div className="max-w-3xl mx-auto space-y-8">
          {/* Hero */}
          <div className="text-center space-y-4 animate-fade-in-up">
            <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-[#3776AB]/15 mb-2">
              <FileUp className="w-8 h-8 text-[#3776AB]" />
            </div>
            <h1 className="text-4xl md:text-5xl font-bold leading-tight">
              <span className="bg-clip-text text-transparent bg-gradient-to-r from-[#3776AB] via-[#5A9FD4] to-white">
                Markdown
              </span>
              <span className="text-white/70"> → Word</span>
            </h1>
            <p className="text-lg text-gray-400 max-w-2xl mx-auto leading-relaxed">
              Upload a <code className="text-[#3776AB] bg-[#3776AB]/10 px-1.5 py-0.5 rounded">.md</code> file
              and download a formatted <code className="text-[#3776AB] bg-[#3776AB]/10 px-1.5 py-0.5 rounded">.docx</code> document.
              Powered by a Python Cloud Function.
            </p>
          </div>

          {/* Upload Card */}
          <Card className="glass-card border-0 animate-fade-in-up animation-delay-100">
            <CardHeader className="pb-3">
              <CardTitle className="text-sm font-medium flex items-center gap-2 text-gray-400">
                <Upload className="w-4 h-4 text-[#3776AB]" />
                Upload Markdown File
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              {/* Drop zone */}
              <div
                onDrop={onDrop}
                onDragOver={onDragOver}
                onDragLeave={onDragLeave}
                onClick={() => inputRef.current?.click()}
                className={`drop-zone cursor-pointer rounded-lg border-2 border-dashed p-10 text-center transition-all ${
                  isDragging
                    ? "border-[#3776AB] bg-[#3776AB]/10"
                    : "border-[#3776AB]/25 hover:border-[#3776AB]/50 hover:bg-[#3776AB]/5"
                }`}
              >
                <input
                  ref={inputRef}
                  type="file"
                  accept=".md,.markdown,.txt"
                  onChange={onInputChange}
                  className="hidden"
                />
                <Upload className="w-10 h-10 mx-auto mb-3 text-[#3776AB]/60" />
                <p className="text-gray-300 text-sm">
                  <span className="text-[#3776AB] font-medium">Click to select</span> or drag &amp; drop
                </p>
                <p className="text-gray-500 text-xs mt-1">
                  Supports .md, .markdown, .txt — up to 10 MB
                </p>
              </div>

              {/* Selected file info */}
              {fileName && (
                <div className="route-card p-4 flex items-center gap-3">
                  <div className="w-10 h-10 rounded-lg bg-[#3776AB]/15 flex items-center justify-center shrink-0">
                    <FileText className="w-5 h-5 text-[#3776AB]" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm text-gray-200 truncate">{fileName}</p>
                    <p className="text-xs text-gray-500">
                      {markdownContent.length.toLocaleString()} characters
                    </p>
                  </div>
                  <button
                    onClick={handleReset}
                    className="text-xs text-gray-500 hover:text-red-400 transition-colors cursor-pointer"
                  >
                    Remove
                  </button>
                </div>
              )}

              {/* Error message */}
              {status === "error" && errorMsg && (
                <div className="flex items-start gap-2 p-3 rounded-lg bg-red-500/10 border border-red-500/30">
                  <AlertCircle className="w-4 h-4 text-red-400 shrink-0 mt-0.5" />
                  <p className="text-sm text-red-300">{errorMsg}</p>
                </div>
              )}

              {/* Action buttons */}
              <div className="flex flex-col sm:flex-row gap-3">
                <Button
                  onClick={handleConvert}
                  disabled={!markdownContent.trim() || status === "converting"}
                  className="btn-primary rounded-lg cursor-pointer flex-1"
                >
                  {status === "converting" ? (
                    <>
                      <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                      Converting…
                    </>
                  ) : (
                    <>
                      <FileText className="w-4 h-4 mr-2" />
                      Convert to DOCX
                    </>
                  )}
                </Button>

                {status === "done" && docxUrl && (
                  <a href={docxUrl} download={downloadName} className="flex-1">
                    <Button className="btn-primary rounded-lg cursor-pointer w-full">
                      <Download className="w-4 h-4 mr-2" />
                      Download .docx
                    </Button>
                  </a>
                )}
              </div>

              {/* Success indicator */}
              {status === "done" && (
                <div className="flex items-center gap-2 p-3 rounded-lg bg-green-500/10 border border-green-500/30">
                  <CheckCircle2 className="w-4 h-4 text-green-400 shrink-0" />
                  <p className="text-sm text-green-300">
                    Conversion complete! Click the download button to save your file.
                  </p>
                </div>
              )}
            </CardContent>
          </Card>

          {/* Supported features */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 animate-fade-in-up animation-delay-200">
            {[
              "Headings (H1–H6)",
              "Bold / Italic / Strikethrough",
              "Inline & block code",
              "Ordered & unordered lists",
              "Blockquotes",
              "Tables",
              "Hyperlinks",
              "Horizontal rules",
            ].map((feat) => (
              <div
                key={feat}
                className="route-card px-3 py-2 text-xs text-gray-400 text-center"
              >
                {feat}
              </div>
            ))}
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="footer-border relative z-10 mt-16">
        <div className="container mx-auto px-6 py-8">
          <div className="flex items-center justify-center gap-2 text-gray-500">
            <span>Powered by</span>
            <a
              href="https://pages.edgeone.ai"
              target="_blank"
              rel="noopener noreferrer"
              className="text-gray-400 hover:text-[#3776AB] transition-colors flex items-center gap-1"
            >
              <img src="/eo-logo-blue.svg" alt="EdgeOne" width={16} height={16} />
              EdgeOne Pages
            </a>
          </div>
        </div>
      </footer>
    </div>
  )
}
