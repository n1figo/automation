'use client'

import { useState, useRef } from 'react'
import { Upload, FileText } from 'lucide-react'
import { Progress } from '@/components/ui/progress'
import { cn } from '@/lib/utils'

interface ProcessingResult {
  type: 'log' | 'progress' | 'result'
  message?: string
  value?: number
  file?: string
}

export default function PDFProcessor() {
  const [file, setFile] = useState<File | null>(null)
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState(0)
  const [logs, setLogs] = useState<string[]>([])
  const [error, setError] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const addLog = (message: string) => {
    setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${message}`])
  }

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0]
    if (selectedFile?.type === 'application/pdf') {
      setFile(selectedFile)
      setError(null)
      addLog(`파일 선택됨: ${selectedFile.name}`)
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
      addLog(`파일 드롭됨: ${droppedFile.name}`)
    } else {
      setError('PDF 파일만 업로드 가능합니다.')
    }
  }

  const handleDragOver = (event: React.DragEvent) => {
    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
  }

  const handleDragEnter = (event: React.DragEvent) => {
    event.preventDefault()
  }

  const handleDragLeave = (event: React.DragEvent) => {
    event.preventDefault()
  }

  const processFile = async () => {
    if (!file) return

    setLoading(true)
    setProgress(0)
    setLogs([])
    addLog('파일 처리 시작...')

    const formData = new FormData()
    formData.append('file', file)

    try {
      const response = await fetch('http://localhost:8000/process-pdf', {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        throw new Error('서버 처리 중 오류가 발생했습니다.')
      }

      const reader = response.body?.getReader()
      if (!reader) {
        throw new Error('Response stream을 읽을 수 없습니다.')
      }

      while (true) {
        const { done, value } = await reader.read()
        if (done) break

        // 서버로부터 받은 데이터 처리
        const text = new TextDecoder().decode(value)
        const lines = text.split('\n').filter(line => line.trim())

        for (const line of lines) {
          try {
            const data: ProcessingResult = JSON.parse(line)
            
            switch (data.type) {
              case 'log':
                if (data.message) addLog(data.message)
                break
              case 'progress':
                if (data.value !== undefined) setProgress(data.value)
                break
              case 'result':
                if (data.file) {
                  // 결과 파일 다운로드
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
                  document.body.removeChild(a)
                  addLog('파일 다운로드 완료')
                }
                break
            }
          } catch (e) {
            console.error('Stream parsing error:', e)
          }
        }
      }

      addLog('처리가 완료되었습니다.')
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : '알 수 없는 오류가 발생했습니다.'
      setError(errorMessage)
      addLog(`오류 발생: ${errorMessage}`)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="w-full max-w-4xl mx-auto p-6 space-y-6">
      <h1 className="text-2xl font-bold text-center">
        보험 약관 PDF 처리 서비스
      </h1>

      {error && (
        <div className="bg-red-50 border-l-4 border-red-500 p-4">
          <p className="text-red-700">{error}</p>
        </div>
      )}

      <div className="flex flex-col items-center gap-4">
        <div
          className={cn(
            "w-full max-w-md border-2 border-dashed rounded-lg p-6 text-center",
            "transition-colors duration-200",
            "hover:bg-gray-50"
          )}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragEnter={handleDragEnter}
          onDragLeave={handleDragLeave}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".pdf"
            className="hidden"
          />
          <Upload className="mx-auto h-12 w-12 text-gray-400" />
          <p className="mt-2 text-sm text-gray-600">
            PDF 파일을 드래그하거나 클릭하여 업로드하세요
          </p>
          <button
            onClick={() => fileInputRef.current?.click()}
            className="mt-4 px-4 py-2 text-sm text-blue-600 border border-blue-600 rounded-md hover:bg-blue-50"
            type="button"
          >
            파일 선택
          </button>
        </div>

        {file && (
          <div className="w-full max-w-md">
            <div className="flex items-center gap-2 text-sm text-gray-600 mb-2">
              <FileText className="h-4 w-4" />
              <span>{file.name}</span>
            </div>
            <button
              onClick={processFile}
              disabled={loading}
              className={cn(
                "w-full px-4 py-2 rounded-md text-white transition-colors",
                loading 
                  ? "bg-blue-400 cursor-not-allowed" 
                  : "bg-blue-600 hover:bg-blue-700"
              )}
            >
              {loading ? '처리중...' : '파일 처리 시작'}
            </button>
          </div>
        )}
      </div>

      {loading && (
        <div className="w-full max-w-md mx-auto">
          <Progress value={progress} className="h-2" />
          <p className="text-center text-sm text-gray-600 mt-2">
            {progress}% 완료
          </p>
        </div>
      )}

      {logs.length > 0 && (
        <div className="w-full max-w-4xl mx-auto mt-8">
          <h2 className="text-lg font-semibold mb-4">처리 로그</h2>
          <div className="bg-black text-green-400 p-4 rounded-lg h-64 overflow-y-auto font-mono text-sm">
            {logs.map((log, i) => (
              <div key={i} className="mb-1">{log}</div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}