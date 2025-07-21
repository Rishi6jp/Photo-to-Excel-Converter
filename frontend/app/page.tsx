'use client'

import { useState } from 'react'

export default function Home() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null)
  const [loading, setLoading] = useState(false)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file) {
      setSelectedFile(file)
    }
  }

  const handleUpload = async () => {
    if (!selectedFile) return alert("Please select an image first.")

    const formData = new FormData()
    formData.append('file', selectedFile)

    try {
      setLoading(true)
      const response = await fetch('https://photo-to-excel-converter.onrender.com/upload-image', {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        const errorText = await response.text()
        throw new Error(`Upload failed: ${errorText}`)
      }

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)

      const a = document.createElement('a')
      a.href = url
      a.download = 'pixel_art.xlsx'
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)
    } catch (err) {
      console.error("Error uploading file:", err)
      alert('Something went wrong while uploading!')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-6">
      <h1 className="text-3xl font-bold mb-6">🧩 Pixel Art Excel Generator</h1>
      
      <input
        type="file"
        accept="image/*"
        onChange={handleFileChange}
        className="mb-4"
      />

      <button
        onClick={handleUpload}
        className="bg-black text-white px-4 py-2 rounded hover:bg-gray-800 disabled:opacity-50"
        disabled={loading}
      >
        {loading ? 'Processing...' : 'Upload & Download Excel'}
      </button>
    </div>
  )
}
