import dynamic from 'next/dynamic'

const PDFProcessor = dynamic(() => import('@/components/PDFProcessor'), {
  ssr: false
})

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <PDFProcessor />
    </main>
  )
}