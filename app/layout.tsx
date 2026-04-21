import type { Metadata } from 'next'

import './globals.css'

export const metadata: Metadata = {
  title: 'Atlas Sheet Viewer',
  description:
    'Server-backed Excel viewer built for large XLSX files and Vercel deployment.',
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}

