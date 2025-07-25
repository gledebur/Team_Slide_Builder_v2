"use client"

import { useState } from "react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { Loader2, Download, Users, FileText } from "lucide-react"

// Dynamic backend URL detection
const getBackendUrl = () => {
  if (typeof window !== 'undefined') {
    // Check if we're in a Codespace environment
    const hostname = window.location.hostname
    if (hostname.includes('.app.github.dev')) {
      // Extract the codespace identifier and construct backend URL
      const parts = hostname.split('.')
      if (parts.length >= 4) {
        const baseUrl = parts.slice(0, -3).join('.') // Remove 'app.github.dev'
        return `https://${baseUrl}-5000.app.github.dev`
      }
    }
    
    // Check if we're running on a different port (e.g., development)
    if (hostname === 'localhost' || hostname === '127.0.0.1') {
      return 'http://localhost:5000'
    }
    
    // Default fallback - try same host with port 5000
    const protocol = window.location.protocol
    return `${protocol}//${hostname}:5000`
  }
  
  // Server-side rendering fallback
  return 'http://localhost:5000'
}

export default function TeamSlideGenerator() {
  const [consultantNames, setConsultantNames] = useState(["", "", "", ""])
  const [isGenerating, setIsGenerating] = useState(false)
  const [generationComplete, setGenerationComplete] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  const handleNameChange = (index: number, value: string) => {
    const newNames = [...consultantNames]
    newNames[index] = value
    setConsultantNames(newNames)
  }

  const handleGenerate = async () => {
    // Validate input
    const filledNames = consultantNames.filter((name) => name.trim() !== "")
    if (filledNames.length !== 4) {
      setError("Please enter exactly 4 consultant names.")
      return
    }

    setIsGenerating(true)
    setError(null)
    setGenerationComplete(false)

    try {
      const backendUrl = getBackendUrl()
      console.log('Using backend URL:', backendUrl) // Debug log
      
      // Call Flask backend to generate team slide
      const response = await fetch(`${backendUrl}/generate`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          consultants: filledNames
        })
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Failed to generate team slide')
      }

      // Create blob URL for download
      const blob = await response.blob()
      const downloadUrl = URL.createObjectURL(blob)
      setDownloadUrl(downloadUrl)
      setGenerationComplete(true)
    } catch (err) {
      console.error('Error generating team slide:', err)
      setError(err instanceof Error ? err.message : "Failed to generate team slide. Please check that all CV files exist and try again.")
    } finally {
      setIsGenerating(false)
    }
  }

  const handleDownload = () => {
    if (downloadUrl) {
      const link = document.createElement("a")
      link.href = downloadUrl
      link.download = "Team_Slide_Output.pptx"
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      
      // Clean up the blob URL to free memory
      URL.revokeObjectURL(downloadUrl)
    }
  }

  const canGenerate = consultantNames.every((name) => name.trim() !== "") && !isGenerating

  return (
    <div className="min-h-screen bg-gray-50 py-8 px-4">
      <div className="max-w-2xl mx-auto">
        <div className="text-center mb-8">
          <div className="flex items-center justify-center mb-4">
            <Users className="h-8 w-8 text-purple-600 mr-2" />
            <h1 className="text-3xl font-bold text-gray-900">Team Slide Generator</h1>
          </div>
          <p className="text-gray-600">Automatically create professional team slides from consultant CV files</p>
        </div>

        <Card className="mb-6">
          <CardHeader>
            <CardTitle className="flex items-center">
              <FileText className="h-5 w-5 mr-2 text-purple-600" />
              Consultant Information
            </CardTitle>
            <CardDescription>
              Enter the names of 4 consultants. The tool will search for matching CV files (Firstname_Lastname.pptx).
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            {consultantNames.map((name, index) => (
              <div key={index} className="space-y-2">
                <Label htmlFor={`consultant-${index}`}>Consultant {index + 1}</Label>
                <Input
                  id={`consultant-${index}`}
                  placeholder="e.g., John Smith"
                  value={name}
                  onChange={(e) => handleNameChange(index, e.target.value)}
                  disabled={isGenerating}
                />
              </div>
            ))}
          </CardContent>
        </Card>

        {error && (
          <Alert className="mb-6 border-red-200 bg-red-50">
            <AlertDescription className="text-red-700">{error}</AlertDescription>
          </Alert>
        )}

        {generationComplete && (
          <Alert className="mb-6 border-green-200 bg-green-50">
            <AlertDescription className="text-green-700">
              Team slide generated successfully! Click the download button below to get your PowerPoint file.
            </AlertDescription>
          </Alert>
        )}

        <div className="flex flex-col sm:flex-row gap-4">
          <Button
            onClick={handleGenerate}
            disabled={!canGenerate}
            className="flex-1 bg-purple-600 hover:bg-purple-700"
            size="lg"
          >
            {isGenerating ? (
              <>
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
                Generating Team Slide...
              </>
            ) : (
              "Generate Team Slide"
            )}
          </Button>

          {generationComplete && (
            <Button
              onClick={handleDownload}
              variant="outline"
              size="lg"
              className="flex-1 sm:flex-initial bg-transparent"
            >
              <Download className="h-4 w-4 mr-2" />
              Download PPTX
            </Button>
          )}
        </div>

        <div className="mt-8 p-4 bg-white rounded-lg border border-gray-200">
          <h3 className="font-semibold text-gray-900 mb-2">How it works:</h3>
          <ol className="text-sm text-gray-600 space-y-1 list-decimal list-inside">
            <li>Enter 4 consultant names in the format "Firstname Lastname"</li>
            <li>The tool searches for CV files named "Firstname_Lastname.pptx"</li>
            <li>Information is extracted: photo, name, role, location, and experience</li>
            <li>A new team slide is created in 2x2 layout format</li>
            <li>Download the generated PowerPoint file</li>
          </ol>
        </div>
      </div>
    </div>
  )
}
