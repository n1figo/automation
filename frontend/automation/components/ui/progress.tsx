// components/PDFProcessor.tsx
'use client'

import { useState, useRef } from 'react'
import { Button } from './ui/button'
import { Progress } from './ui/progress'
import { FileUp } from 'lucide-react'

export default function PDFProcessor() {
  const [file, setFile] = useState<File | null>(null)
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState(0)
  const [logs, setLogs] = useState<string[]>([])
  const [error, setError] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0]
    if (selectedFile?.type === 'application/pdf') {
      setFile(selectedFile)
      setError(null)
    } else {
      setError('PDF 파일만 업로드 가능합니다.')
      setFile(null)
    }
  }

  const handleDrop = (event: React.DragEvent) => {
    event.preventDefault()
    const droppedFile = event.dataTransfer.files[0]
    if (droppedFile?.type === 'application/pdf') {
      setFile(droppedFile)
      setError(null)
    } else {
      setError('PDF 파일만 업로드 가능합니다.')
    }
  }

  const handleDragOver = (event: React.DragEvent) => {
    event.preventDefault()
  }

  const processFile = async () => {
    if (!file) return

    setLoading(true)
    setProgress(0)
    setLogs([])

    const formData = new FormData()
    formData.append('file', file)

    try {
      const response = await fetch('http://localhost:8000/process-pdf', {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) throw new Error('처리 중 오류가 발생했습니다.')

      const reader = response.body?.getReader()
      if (!reader) throw new Error('Response reader error')

      while (true) {
        const { done, value } = await reader.read()
        if (done) break

        const text = new TextDecoder().decode(value)
        const lines = text.split('\n').filter(line => line.trim())

        for (const line of lines) {
          try {
            const data = JSON.parse(line)
            if (data.type === 'log') {
              setLogs(prev => [...prev, data.message])
            } else if (data.type === 'progress') {
              setProgress(data.value)
            } else if (data.type === 'result') {
              // 엑셀 파일 다운로드
              const blob = new Blob(
                [new Uint8Array([...atob(data.file)].map(c => c.charCodeAt(0)))],
                { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
              )
              const url = window.URL.createObjectURL(blob)
              const a = document.createElement('a')
              a.href = url
              a.download = '보험특약표.xlsx'
              document.body.appendChild(a)
              a.click()
              window.URL.revokeObjectURL(url)
            }
          } catch (e) {
            console.error('Stream parsing error:', e)
          }
        }
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '알 수 없는 오류가 발생했습니다.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="w-full max-w-4xl mx-auto p-6 space-y-6">
      <h1 className="text-2xl font-bold text-center mb-8">
        보험 약관 PDF 처리 서비스
      </h1>

      {error && (
        <div className="bg-red-50 border-l-4 border-red-500 p-4 mb-4">
          <p className="text-red-700">{error}</p>
        </div>
      )}

      <div className="flex flex-col items-center gap-4">
        <div
          className="w-full max-w-md border-2 border-dashed rounded-lg p-6 text-center"
          onDrop={handleDrop}
          onDragOver={handleDragOver}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".pdf"
            className="hidden"
          />
          <FileUp className="mx-auto h-12 w-12 text-gray-400" />
          <p className="mt-2 text-sm font-semibold">
            PDF 파일을 드래그하거나 클릭하여 업로드하세요
          </p>
          <Button
            onClick={() => fileInputRef.current?.click()}
            variant="outline"
            className="mt-4"
          >
            파일 선택
          </Button>
        </div>

        {file && (
          <div className="w-full max-w-md">
            <p className="text-sm text-gray-500 mb-2">
              선택된 파일: {file.name}
            </p>
            <Button
              onClick={processFile}
              disabled={loading}
              className="w-full"
            >
              {loading ? '처리중...' : '파일 처리 시작'}
            </Button>
          </div>
        )}
      </div>

      {loading && (
        <Progress value={progress} className="w-full max-w-md mx-auto" />
      )}

      {logs.length > 0 && (
        <div className="w-full max-w-4xl mx-auto mt-8">
          <h2 className="text-lg font-semibold mb-4">처리 로그</h2>
          <div className="bg-black text-green-400 p-4 rounded-lg h-96 overflow-y-auto font-mono text-sm">
            {logs.map((log, i) => (
              <div key={i} className="mb-1">
                {log}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}