'use client'

import { useState, useRef, useEffect } from 'react'
import { Upload, FileText, ChevronDown, ChevronUp, X, Download, Filter } from 'lucide-react'
import { Progress } from '@/components/ui/progress'

interface Log {
  timestamp: string
  level: 'INFO' | 'ERROR' | 'SUCCESS'
  message: string
}

interface AnalysisResult {
  section: string
  changes: string[]
  status: string
}

export default function PDFProcessor() {
  const [isHovered, setIsHovered] = useState(false)
  const [file, setFile] = useState<File | null>(null)
  const [showLogs, setShowLogs] = useState(true)
  const [logs, setLogs] = useState<Log[]>([])
  const [uploadProgress, setUploadProgress] = useState(0)
  const [isUploading, setIsUploading] = useState(false)
  const [analysisResults, setAnalysisResults] = useState<AnalysisResult[]>([])
  const [selectedLogLevels, setSelectedLogLevels] = useState<Set<string>>(
    new Set(['INFO', 'ERROR', 'SUCCESS'])
  )

  const fileInputRef = useRef<HTMLInputElement>(null)
  const logsEndRef = useRef<HTMLDivElement>(null)

  // 자동 스크롤
  useEffect(() => {
    if (showLogs && logsEndRef.current) {
      logsEndRef.current.scrollIntoView({ behavior: 'smooth' })
    }
  }, [logs, showLogs])

  const addLog = (level: Log['level'], message: string) => {
    const newLog = {
      timestamp: new Date().toLocaleTimeString(),
      level,
      message
    }
    setLogs(prev => [...prev, newLog])
  }

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0]
    if (selectedFile) {
      if (selectedFile.type !== 'application/pdf') {
        addLog('ERROR', 'PDF 파일만 업로드 가능합니다.')
        return
      }
      if (selectedFile.size > 10 * 1024 * 1024) {
        addLog('ERROR', '파일 크기는 10MB를 초과할 수 없습니다.')
        return
      }
      setFile(selectedFile)
      addLog('INFO', `파일이 선택되었습니다: ${selectedFile.name}`)
    }
  }

  const handleUpload = async () => {
    if (!file) return

    try {
      setIsUploading(true)
      addLog('INFO', '파일 분석을 시작합니다...')

      // FormData 생성
      const formData = new FormData()
      formData.append('file', file)

      // 실제 API 엔드포인트로 변경 필요
      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData,
        onUploadProgress: (progressEvent) => {
          const progress = (progressEvent.loaded / progressEvent.total) * 100
          setUploadProgress(progress)
        },
      })

      if (!response.ok) throw new Error('업로드 실패')

      const result = await response.json()
      setAnalysisResults(result.results)
      addLog('SUCCESS', '분석이 완료되었습니다.')

    } catch (error) {
      addLog('ERROR', `분석 중 오류가 발생했습니다: ${error.message}`)
    } finally {
      setIsUploading(false)
      setUploadProgress(0)
    }
  }

  const handleLogExport = () => {
    const logText = logs
      .map(log => `[${log.timestamp}] [${log.level}] ${log.message}`)
      .join('\n')
    
    const blob = new Blob([logText], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'analysis-logs.txt'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const handleResultsExport = () => {
    const resultsText = JSON.stringify(analysisResults, null, 2)
    const blob = new Blob([resultsText], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'analysis-results.json'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const filteredLogs = logs.filter(log => selectedLogLevels.has(log.level))

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      {/* Header */}
      <div className="max-w-6xl mx-auto mb-8">
        <h1 className="text-2xl font-bold text-gray-900 mb-2">
          KB 손해보험 상품개정 자동화서비스
        </h1>
        <p className="text-gray-600">
          PDF 파일을 업로드하여 상품 개정 내용을 자동으로 분석합니다.
        </p>
      </div>

      <div className="max-w-6xl mx-auto grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="space-y-6">
          {/* Upload Section */}
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileSelect}
            accept=".pdf"
            className="hidden"
          />

          <div 
            onClick={() => fileInputRef.current?.click()}
            className={`
              border-2 border-dashed rounded-lg p-8 bg-white cursor-pointer
              ${isHovered ? 'border-gray-400 bg-gray-50' : 'border-gray-300'} 
              transition-all duration-200
            `}
            onDragOver={(e) => {
              e.preventDefault()
              setIsHovered(true)
            }}
            onDragLeave={() => setIsHovered(false)}
            onDrop={(e) => {
              e.preventDefault()
              setIsHovered(false)
              const droppedFile = e.dataTransfer.files[0]
              if (droppedFile) {
                if (droppedFile.type === 'application/pdf') {
                  setFile(droppedFile)
                  addLog('INFO', `파일이 드롭되었습니다: ${droppedFile.name}`)
                } else {
                  addLog('ERROR', 'PDF 파일만 업로드 가능합니다.')
                }
              }
            }}
          >
            <div className="flex flex-col items-center justify-center gap-4">
              <div className="p-4 bg-gray-100 rounded-full">
                {file ? (
                  <FileText className="w-8 h-8 text-blue-500" />
                ) : (
                  <Upload className="w-8 h-8 text-gray-400" />
                )}
              </div>

              <div className="space-y-2 text-center">
                {file ? (
                  <>
                    <p className="text-sm font-medium">{file.name}</p>
                    <p className="text-xs text-gray-500">
                      {(file.size / 1024 / 1024).toFixed(2)} MB
                    </p>
                    <button
                      onClick={(e) => {
                        e.stopPropagation()
                        setFile(null)
                        addLog('INFO', '파일이 제거되었습니다.')
                      }}
                      className="text-red-500 hover:text-red-600 text-sm flex items-center gap-1 mx-auto"
                    >
                      <X className="w-4 h-4" />
                      파일 제거
                    </button>
                  </>
                ) : (
                  <>
                    <p className="text-sm font-medium">
                      PDF 파일을 드래그하거나 클릭하여 업로드
                    </p>
                    <p className="text-xs text-gray-500">
                      최대 10MB까지 업로드 가능
                    </p>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Upload Progress */}
          {isUploading && (
            <div className="space-y-2">
              <Progress value={uploadProgress} />
              <p className="text-sm text-center text-gray-500">
                {Math.round(uploadProgress)}% 완료
              </p>
            </div>
          )}

          {/* Upload Button */}
          {file && !isUploading && (
            <button
              onClick={handleUpload}
              className="w-full py-2 px-4 bg-blue-500 text-white rounded-lg font-medium hover:bg-blue-600 transition-colors"
            >
              분석 시작
            </button>
          )}
        </div>

        <div className="space-y-6">
          {/* Logs Section */}
          <div className="bg-white rounded-lg shadow">
            <div className="px-4 py-3 flex items-center justify-between bg-gray-50 border-b">
              <button 
                onClick={() => setShowLogs(!showLogs)}
                className="font-medium text-gray-700 flex items-center gap-2"
              >
                <span>처리 로그</span>
                {showLogs ? (
                  <ChevronUp className="w-5 h-5" />
                ) : (
                  <ChevronDown className="w-5 h-5" />
                )}
              </button>
              <div className="flex items-center gap-2">
                <div className="flex items-center gap-2 text-sm">
                  {['INFO', 'ERROR', 'SUCCESS'].map((level) => (
                    <label key={level} className="flex items-center gap-1">
                      <input
                        type="checkbox"
                        checked={selectedLogLevels.has(level)}
                        onChange={(e) => {
                          const newLevels = new Set(selectedLogLevels)
                          if (e.target.checked) {
                            newLevels.add(level)
                          } else {
                            newLevels.delete(level)
                          }
                          setSelectedLogLevels(newLevels)
                        }}
                        className="rounded text-blue-500"
                      />
                      {level}
                    </label>
                  ))}
                </div>
                <button
                  onClick={handleLogExport}
                  className="p-1 hover:bg-gray-100 rounded-full"
                  title="로그 내보내기"
                >
                  <Download className="w-4 h-4" />
                </button>
              </div>
            </div>
            
            {showLogs && (
              <div className="p-4 h-60 overflow-y-auto">
                {filteredLogs.map((log, index) => (
                  <div 
                    key={index}
                    className="py-1 font-mono text-sm"
                  >
                    <span className={`
                      ${log.level === 'INFO' ? 'text-blue-600' : ''}
                      ${log.level === 'ERROR' ? 'text-red-600' : ''}
                      ${log.level === 'SUCCESS' ? 'text-green-600' : ''}
                    `}>
                      [{log.timestamp}] [{log.level}] {log.message}
                    </span>
                  </div>
                ))}
                <div ref={logsEndRef} />
              </div>
            )}
          </div>

          {/* Analysis Results */}
          {analysisResults.length > 0 && (
            <div className="bg-white rounded-lg shadow">
              <div className="px-4 py-3 flex items-center justify-between bg-gray-50 border-b">
                <span className="font-medium text-gray-700">분석 결과</span>
                <button
                  onClick={handleResultsExport}
                  className="p-1 hover:bg-gray-100 rounded-full"
                  title="결과 내보내기"
                >
                  <Download className="w-4 h-4" />
                </button>
              </div>
              <div className="p-4 max-h-60 overflow-y-auto">
                <table className="min-w-full">
                  <thead>
                    <tr>
                      <th className="text-left text-sm font-medium text-gray-500 pb-2">섹션</th>
                      <th className="text-left text-sm font-medium text-gray-500 pb-2">변경사항</th>
                      <th className="text-left text-sm font-medium text-gray-500 pb-2">상태</th>
                    </tr>
                  </thead>
                  <tbody>
                    {analysisResults.map((result, index) => (
                      <tr key={index} className="border-t">
                        <td className="py-2 text-sm">{result.section}</td>
                        <td className="py-2 text-sm">
                          <ul className="list-disc list-inside">
                            {result.changes.map((change, i) => (
                              <li key={i}>{change}</li>
                            ))}
                          </ul>
                        </td>
                        <td className="py-2 text-sm">
                          <span className={`
                            px-2 py-1 rounded-full text-xs
                            ${result.status === 'completed' ? 'bg-green-100 text-green-800' : ''}
                            ${result.status === 'pending' ? 'bg-yellow-100 text-yellow-800' : ''}
                            ${result.status === 'error' ? 'bg-red-100 text-red-800' : ''}
                          `}>
                            {result.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}