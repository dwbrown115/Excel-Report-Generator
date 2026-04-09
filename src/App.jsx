import { useState, useCallback } from 'react'
import readXlsxFile from 'read-excel-file/browser'
import './App.css'

function App() {
  const [uniqueValues, setUniqueValues] = useState([])
  const [error, setError] = useState(null)
  const [fileName, setFileName] = useState(null)
  const [loading, setLoading] = useState(false)

  const handleFileChange = useCallback(async (e) => {
    const file = e.target.files[0]
    if (!file) return

    setFileName(file.name)
    setError(null)
    setUniqueValues([])
    setLoading(true)

    try {
      const rows = await readXlsxFile(file)
      // Collect every non-empty value from the first column (column A)
      const seen = new Set()
      const unique = []
      for (const row of rows) {
        const cell = row[0]
        if (cell !== null && cell !== undefined && cell !== '') {
          const str = String(cell)
          if (!seen.has(str)) {
            seen.add(str)
            unique.push(str)
          }
        }
      }
      setUniqueValues(unique)
    } catch (err) {
      setError('Failed to parse the file. Please upload a valid .xlsx file.')
      console.error(err)
    } finally {
      setLoading(false)
    }
  }, [])

  return (
    <div className="app">
      <header className="app-header">
        <h1>Excel Report Generator</h1>
        <p className="subtitle">Upload an Excel file to view unique Column A values</p>
      </header>

      <main className="app-main">
        <section className="upload-section">
          <label className="file-label" htmlFor="excel-input">
            <span className="file-icon">📂</span>
            <span>{fileName ? fileName : 'Choose an Excel file (.xlsx)'}</span>
            <input
              id="excel-input"
              type="file"
              accept=".xlsx"
              onChange={handleFileChange}
              className="file-input"
            />
          </label>
        </section>

        {loading && <p className="status">Processing file…</p>}

        {error && <p className="error">{error}</p>}

        {!loading && uniqueValues.length > 0 && (
          <section className="results-section">
            <h2>
              Unique Column A Values{' '}
              <span className="count-badge">{uniqueValues.length}</span>
            </h2>
            <ol className="values-list">
              {uniqueValues.map((value, index) => (
                <li key={index}>{value}</li>
              ))}
            </ol>
          </section>
        )}

        {!loading && fileName && uniqueValues.length === 0 && !error && (
          <p className="status">No values found in Column A.</p>
        )}
      </main>
    </div>
  )
}

export default App
