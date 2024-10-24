'use client'

import { useState, useRef } from 'react'
import { Upload, Download, FileText } from 'lucide-react'
import { Progress } from '@/components/ui/progress'

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

  const addLog = (message: string) => {
    setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${message}`])
  }

  const processFile = async () => {
    if (!file) return

    setLoading(true)
    setProgress(0)
    setLogs([])
    addLog('처리 시작...')

    // 여기서는 테스트를 위해 진행 상황을 시뮬레이션합니다
    // 실제로는 백엔드 API와 통신하게 됩니다
    try {
      for (let i = 0; i <= 100; i += 10) {
        setProgress(i)
        addLog(`진행률: ${i}%`)
        await new Promise(r => setTimeout(r, 500))
      }

      // 처리 완료 시 다운로드 시뮬레이션
      addLog('처리 완료! 파일 다운로드를 시작합니다.')
      
      // 실제 구현에서는 백엔드에서 받은 파일을 다운로드합니다
      setTimeout(() => {
        const link = document.createElement('a')
        link.href = 'data:text/plain;charset=utf-8,처리완료'
        link.download = '처리결과.xlsx'
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
      }, 1000)

    } catch (err) {
      setError(err instanceof Error ? err.message : '처리 중 오류가 발생했습니다.')
      addLog('오류 발생!')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="w-full max-w-4xl mx-auto p-6 space-y-6">
      <h1 className="text-2xl font-bold text-center">
        PDF 문서 처리 서비스
      </h1>

      {error && (
        <div className="bg-red-50 border-l-4 border-red-500 p-4">
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
          <Upload className="mx-auto h-12 w-12 text-gray-400" />
          <p className="mt-2 text-sm text-gray-600">
            PDF 파일을 드래그하거나 클릭하여 업로드하세요
          </p>
          <button
            onClick={() => fileInputRef.current?.click()}
            className="mt-4 px-4 py-2 text-sm text-blue-600 border border-blue-600 rounded-md hover:bg-blue-50"
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
              className={`w-full px-4 py-2 rounded-md text-white ${
                loading 
                  ? 'bg-blue-400 cursor-not-allowed' 
                  : 'bg-blue-600 hover:bg-blue-700'
              }`}
            >
              {loading ? '처리중...' : '파일 처리 시작'}
            </button>
          </div>
        )}
      </div>

      {loading && (
        <div className="w-full max-w-md mx-auto">
          <Progress value={progress} className="h-2 bg-gray-200" />
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