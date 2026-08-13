import React, { useState, useRef } from 'react';
import { 
  FileSpreadsheet, 
  UploadCloud, 
  Trash2, 
  Play, 
  CheckCircle2, 
  Download, 
  RefreshCw, 
  AlertCircle, 
  Info, 
  Sparkles,
  Check
} from 'lucide-react';

function App() {
  const [file, setFile] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(false);

  // Excel Metadata
  const [students, setStudents] = useState([]);
  const [subjects, setSubjects] = useState([]);

  // User input & LLM result
  const [prompt, setPrompt] = useState('');
  const [updates, setUpdates] = useState([]);
  const [processed, setProcessed] = useState(false);

  const fileInputRef = useRef(null);

  // 1. Upload & Analyze Excel
  const uploadAndAnalyze = async (selectedFile) => {
    if (!selectedFile) return;
    
    // Check if it's an Excel file
    if (!selectedFile.name.endsWith('.xlsx')) {
      setError("Only Excel (.xlsx) files are supported.");
      return;
    }

    setFile(selectedFile);
    setLoading(true);
    setError(null);
    setSuccess(false);
    setStudents([]);
    setSubjects([]);
    setUpdates([]);
    setProcessed(false);

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Failed to analyze Excel file.");
      }

      setStudents(data.students || []);
      setSubjects(data.subjects || []);
    } catch (err) {
      console.error(err);
      setError(err.message || "An error occurred while uploading the file.");
      setFile(null);
    } finally {
      setLoading(false);
    }
  };

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    uploadAndAnalyze(selectedFile);
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      uploadAndAnalyze(e.dataTransfer.files[0]);
    }
  };

  const handleRemoveFile = () => {
    setFile(null);
    setStudents([]);
    setSubjects([]);
    setUpdates([]);
    setPrompt('');
    setProcessed(false);
    setSuccess(false);
    setError(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // 2. Process Prompt with LLM
  const handleProcessPrompt = async (e) => {
    e.preventDefault();
    if (!file) {
      setError("Please upload an Excel file first.");
      return;
    }
    if (!prompt.trim()) {
      setError("Please type a instruction prompt.");
      return;
    }

    setLoading(true);
    setError(null);
    setUpdates([]);
    setProcessed(false);

    try {
      const response = await fetch('/api/process-prompt', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          prompt,
          students,
          subjects
        }),
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Failed to process prompt.");
      }

      setUpdates(data.updates || []);
      setProcessed(true);
    } catch (err) {
      console.error(err);
      setError(err.message || "Failed to extract updates using AI.");
    } finally {
      setLoading(false);
    }
  };

  // Remove individual update row
  const handleDeleteUpdate = (indexToDelete) => {
    setUpdates(prev => prev.filter((_, idx) => idx !== indexToDelete));
  };

  // 3. Confirm and Download Updated Excel
  const handleConfirmAndDownload = async () => {
    if (!file) return;
    if (updates.length === 0) {
      setError("No updates to apply.");
      return;
    }

    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('updates', JSON.stringify(updates));

    try {
      const response = await fetch('/api/confirm', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || "Failed to apply updates and download sheet.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      // Append "_updated" to the original file name
      const originalName = file.name.substring(0, file.name.lastIndexOf('.'));
      a.download = `${originalName}_updated.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      setSuccess(true);
    } catch (err) {
      console.error(err);
      setError(err.message || "Failed to download updated file.");
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    handleRemoveFile();
  };

  return (
    <div className="app-container">
      {/* Header */}
      <header className="app-header">
        <div className="logo-icon">
          <FileSpreadsheet size={32} />
        </div>
        <h1 className="app-title">AI Excel Result Assistant</h1>
        <p className="app-subtitle">
          Upload a spreadsheet and describe the student marks in natural language to update it using AI.
        </p>
      </header>

      {/* Error Alert */}
      {error && (
        <div className="alert alert-danger glass-card">
          <AlertCircle size={20} className="alert-icon" />
          <div>
            <strong>Error:</strong> {error}
          </div>
        </div>
      )}

      {/* Main App Workflow */}
      {success ? (
        <div className="glass-card success-view">
          <div className="success-icon">
            <CheckCircle2 size={48} />
          </div>
          <h2>Excel Sheet Updated!</h2>
          <p style={{ color: '#94a3b8', maxWidth: '500px', margin: '0 auto' }}>
            Your updated file has been processed and downloaded. All student totals, percentages, and positions/ranks have been recalculated successfully.
          </p>
          <div style={{ display: 'flex', gap: '1rem', width: '100%', maxWidth: '300px', marginTop: '1rem' }}>
            <button className="btn btn-primary" onClick={handleReset}>
              <RefreshCw size={18} />
              Process Another Sheet
            </button>
          </div>
        </div>
      ) : (
        <div className="dashboard-grid">
          {/* Sidebar Panel - Upload & Metadata */}
          <div className="sidebar-panel">
            <div className="glass-card">
              <h2 className="section-title">
                <UploadCloud size={20} />
                1. Upload Sheet
              </h2>
              
              {!file ? (
                <div 
                  className={`upload-zone ${dragActive ? 'active' : ''}`}
                  onDragEnter={handleDrag}
                  onDragOver={handleDrag}
                  onDragLeave={handleDrag}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                >
                  <input 
                    type="file" 
                    ref={fileInputRef}
                    onChange={handleFileChange}
                    accept=".xlsx"
                    style={{ display: 'none' }}
                  />
                  <UploadCloud size={40} className="upload-icon" />
                  <p className="upload-text">Click or drag & drop</p>
                  <p className="upload-hint">Only Excel (.xlsx) files are supported</p>
                </div>
              ) : (
                <div className="file-badge">
                  <div className="file-info">
                    <FileSpreadsheet size={20} style={{ color: '#10b981' }} />
                    <span className="file-name" title={file.name}>{file.name}</span>
                  </div>
                  <button className="btn-remove-file" onClick={handleRemoveFile} title="Remove file">
                    <Trash2 size={16} />
                  </button>
                </div>
              )}
            </div>

            {/* Display Detected Metadata */}
            {file && (students.length > 0 || subjects.length > 0) && (
              <div className="glass-card">
                <h2 className="section-title">
                  <Info size={20} />
                  Detected Data
                </h2>

                <div className="metadata-group">
                  <div className="metadata-label">
                    <span>Students</span>
                    <span className="metadata-count">{students.length}</span>
                  </div>
                  <div className="badge-grid">
                    {students.map((student, i) => (
                      <span key={i} className="badge badge-student">{student}</span>
                    ))}
                  </div>
                </div>

                <div className="metadata-group" style={{ marginBottom: 0 }}>
                  <div className="metadata-label">
                    <span>Subjects</span>
                    <span className="metadata-count">{subjects.length}</span>
                  </div>
                  <div className="badge-grid">
                    {subjects.map((subj, i) => (
                      <span key={i} className="badge badge-subject">{subj}</span>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Main Action Panel */}
          <div className="main-panel">
            {/* Step 2 - AI Input Prompt */}
            <div className="glass-card" style={{ opacity: file ? 1 : 0.5, pointerEvents: file ? 'auto' : 'none' }}>
              <h2 className="section-title">
                <Sparkles size={20} style={{ color: '#a855f7' }} />
                2. AI Instruction
              </h2>
              
              <form onSubmit={handleProcessPrompt} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                <textarea
                  className="prompt-textarea"
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  placeholder={
                    file 
                      ? `Type which marks to enter, e.g.:\n"${students[0] || 'Ali'} got 85 in ${subjects[0] || 'Math'} and 78 in ${subjects[1] || 'Urdu'}."`
                      : "Upload an Excel sheet first..."
                  }
                  disabled={!file || loading}
                />
                
                <div className="instructions-card">
                  <strong>How to prompt:</strong>
                  <ul>
                    <li>You can enter marks for multiple students and subjects in one go.</li>
                    <li>The AI will automatically map names to the exact casing in your Excel file.</li>
                    <li>Example: <em>"Ahmed scored 90 in English, while Ali got 82 in Physics."</em></li>
                  </ul>
                </div>

                <button 
                  type="submit" 
                  className="btn btn-primary" 
                  disabled={loading || !prompt.trim() || !file}
                >
                  {loading && !processed ? (
                    <svg className="spinner" viewBox="0 0 50 50" width="18" height="18">
                      <circle className="path" cx="25" cy="25" r="20" fill="none" strokeWidth="4"></circle>
                    </svg>
                  ) : (
                    <Play size={16} />
                  )}
                  Process with AI
                </button>
              </form>
            </div>

            {/* Step 3 - Extracted updates table preview */}
            {processed && (
              <div className="glass-card">
                <h2 className="section-title">
                  <CheckCircle2 size={20} style={{ color: '#10b981' }} />
                  3. Verify Extracted Changes
                </h2>

                {updates.length === 0 ? (
                  <div style={{ textAlign: 'center', padding: '1.5rem', color: '#94a3b8' }}>
                    No updates were extracted from your text. Please try typing a different prompt.
                  </div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                    <p style={{ color: '#94a3b8', fontSize: '0.9rem', margin: 0 }}>
                      The AI extracted the following updates. Please verify before writing to Excel:
                    </p>

                    <div className="table-container">
                      <table className="preview-table">
                        <thead>
                          <tr>
                            <th>Student</th>
                            <th>Subject</th>
                            <th>Marks</th>
                            <th style={{ width: '50px', textAlign: 'center' }}>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {updates.map((update, idx) => (
                            <tr key={idx}>
                              <td>
                                <strong style={{ color: '#f1f5f9' }}>{update.student}</strong>
                              </td>
                              <td>
                                <span className="badge badge-subject">{update.subject}</span>
                              </td>
                              <td>
                                <strong style={{ color: '#10b981' }}>{update.marks}</strong>
                              </td>
                              <td style={{ textAlign: 'center' }}>
                                <button 
                                  className="btn-delete-row"
                                  onClick={() => handleDeleteUpdate(idx)}
                                  title="Delete update"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <button 
                      className="btn btn-success" 
                      onClick={handleConfirmAndDownload} 
                      disabled={loading || updates.length === 0}
                    >
                      {loading ? (
                        <svg className="spinner" viewBox="0 0 50 50" width="18" height="18">
                          <circle className="path" cx="25" cy="25" r="20" fill="none" strokeWidth="4"></circle>
                        </svg>
                      ) : (
                        <Download size={18} />
                      )}
                      Apply & Download Excel
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
