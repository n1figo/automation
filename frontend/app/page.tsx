'use client'

import { PDFProcessor } from '@/components/PDFProcessor'

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <PDFProcessor />
    </main>
  )
}

