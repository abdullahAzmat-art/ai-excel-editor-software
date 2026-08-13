import React, { useState, useRef } from 'react';
import {
  FileSpreadsheet, UploadCloud, Trash2, Play, CheckCircle2,
  Download, AlertCircle, Sparkles, UserCheck
} from 'lucide-react';

function App() {
  const [file, setFile] = useState(null);
  const [dragActive, setDragActive] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [successMsg, setSuccessMsg] = useState(null);

  // Session state
  const [sessionId, setSessionId] = useState(null);
  const [students, setStudents] = useState([]);
  const [subjects, setSubjects] = useState([]);
  const [updatedCount, setUpdatedCount] = useState(0);
  const [updatedStudents, setUpdatedStudents] = useState([]);

  // mode: 'upload' | 'edit' | 'done'
  const [mode, setMode] = useState('upload');

  // Prompt and extracted updates
  const [prompt, setPrompt] = useState('');
  const [updates, setUpdates] = useState([]);

  const fileInputRef = useRef(null);

  // ─── 1. Upload & Analyze ───────────────────────────────────────────────────
  const uploadAndAnalyze = async (selectedFile) => {
    if (!selectedFile) return;
    if (!selectedFile.name.endsWith('.xlsx')) {
      setError("Only Excel (.xlsx) files are supported.");
      return;
    }

    setFile(selectedFile);
    setLoading(true);
    setError(null);
    setSuccessMsg(null);

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
      const response = await fetch('/api/analyze', { method: 'POST', body: formData });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Failed to analyze Excel file.");

      setSessionId(data.sessionId);
      setStudents(data.students || []);
      setSubjects(data.subjects || []);
      setUpdatedCount(0);
      setUpdatedStudents([]);
      setMode('edit');
      setSuccessMsg(`✅ Template loaded! Found ${data.students.length} students and ${data.subjects.length} subjects.`);
    } catch (err) {
      setError(err.message || "An error occurred while uploading the file.");
      setFile(null);
    } finally {
      setLoading(false);
    }
  };

  const handleFileChange = (e) => uploadAndAnalyze(e.target.files[0]);
  const handleDrag = (e) => {
    e.preventDefault(); e.stopPropagation();
    setDragActive(e.type === 'dragenter' || e.type === 'dragover');
  };
  const handleDrop = (e) => {
    e.preventDefault(); e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files?.[0]) uploadAndAnalyze(e.dataTransfer.files[0]);
  };

  const handleReset = () => {
    setFile(null); setSessionId(null); setStudents([]); setSubjects([]);
    setUpdates([]); setPrompt(''); setUpdatedCount(0); setUpdatedStudents([]);
    setMode('upload'); setSuccessMsg(null); setError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // ─── 2. Process Prompt via LLM ─────────────────────────────────────────────
  const handleProcessPrompt = async (e) => {
    e.preventDefault();
    if (!prompt.trim()) return setError("Please type a prompt.");
    setLoading(true); setError(null); setSuccessMsg(null);

    try {
      const response = await fetch('/api/process-prompt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionId, prompt }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Failed to process prompt.");
      setUpdates(data.updates || []);
    } catch (err) {
      setError(err.message || "AI extraction failed.");
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteUpdate = (idx) => setUpdates(prev => prev.filter((_, i) => i !== idx));

  // ─── 3. Apply Updates to Workbook ──────────────────────────────────────────
  const handleApplyUpdates = async () => {
    if (updates.length === 0) return setError("No updates to apply.");
    setLoading(true); setError(null);

    try {
      const response = await fetch('/api/apply-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionId, updates }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Failed to apply updates.");

      setUpdatedCount(data.updatedCount);
      setUpdatedStudents(data.updatedStudents || []);
      setUpdates([]);
      setPrompt('');
      setSuccessMsg(`✅ Marks saved! ${data.updatedCount} of ${students.length} students updated so far.`);
    } catch (err) {
      setError(err.message || "Failed to save marks.");
    } finally {
      setLoading(false);
    }
  };

  // ─── 4. Download ───────────────────────────────────────────────────────────
  const handleDownload = async () => {
    try {
      const response = await fetch(`/api/download/${sessionId}`);
      if (!response.ok) throw new Error("Download failed.");
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${file.name.replace('.xlsx', '')}_updated.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message || "Download failed.");
    }
  };

  const progressPct = students.length > 0 ? Math.round((updatedCount / students.length) * 100) : 0;

  return (
    <div className="app-container">
      <header className="app-header">
        <div className="logo-icon"><FileSpreadsheet size={32} /></div>
        <h1 className="app-title">AI Excel Result Assistant</h1>
        <p className="app-subtitle">Fill student marks into your transcript template using natural language</p>
      </header>

      {error && (
        <div className="alert alert-danger glass-card" style={{ padding: '1rem', marginBottom: '1rem', display: 'flex', gap: '0.75rem', alignItems: 'flex-start' }}>
          <AlertCircle size={20} style={{ flexShrink: 0, color: '#f87171', marginTop: '2px' }} />
          <div><strong>Error:</strong> {error}</div>
        </div>
      )}
      {successMsg && (
        <div className="alert alert-success glass-card" style={{ padding: '1rem', marginBottom: '1rem', display: 'flex', gap: '0.75rem', alignItems: 'flex-start' }}>
          <CheckCircle2 size={20} style={{ flexShrink: 0, color: '#34d399', marginTop: '2px' }} />
          <div>{successMsg}</div>
        </div>
      )}

      <div className="dashboard-grid">
        {/* ── Sidebar ── */}
        <div className="sidebar-panel">

          {/* Upload Card */}
          <div className="glass-card">
            <h2 className="section-title"><UploadCloud size={20} /> Template File</h2>

            {!file ? (
              <div
                className={`upload-zone ${dragActive ? 'active' : ''}`}
                onDragEnter={handleDrag} onDragOver={handleDrag}
                onDragLeave={handleDrag} onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <input ref={fileInputRef} type="file" accept=".xlsx"
                  onChange={handleFileChange} style={{ display: 'none' }} />
                <UploadCloud size={40} className="upload-icon" />
                <p className="upload-text">Drop your Excel template here</p>
                <p style={{ fontSize: '0.8rem', color: '#64748b' }}>or click to browse</p>
              </div>
            ) : (
              <div>
                <div className="file-badge" style={{ marginTop: 0 }}>
                  <div className="file-info">
                    <FileSpreadsheet size={20} style={{ color: '#10b981' }} />
                    <span className="file-name" title={file.name}>{file.name}</span>
                  </div>
                  <button className="btn-remove-file" onClick={handleReset}><Trash2 size={16} /></button>
                </div>

                {/* Progress */}
                {mode === 'edit' && (
                  <div style={{ marginTop: '1.25rem' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.5rem' }}>
                      <span style={{ fontSize: '0.85rem', color: '#94a3b8' }}>Students Updated</span>
                      <span style={{ fontSize: '0.85rem', fontWeight: 700, color: '#10b981' }}>
                        {updatedCount} / {students.length}
                      </span>
                    </div>
                    <div style={{ height: '6px', background: 'rgba(255,255,255,0.08)', borderRadius: '99px', overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: `${progressPct}%`, background: 'linear-gradient(90deg, #6366f1, #10b981)', borderRadius: '99px', transition: 'width 0.4s ease' }} />
                    </div>
                    <p style={{ fontSize: '0.75rem', color: '#475569', marginTop: '0.4rem' }}>{progressPct}% complete</p>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Updated Students List */}
          {updatedStudents.length > 0 && (
            <div className="glass-card">
              <h2 className="section-title"><UserCheck size={18} /> Updated Students</h2>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.4rem', maxHeight: '240px', overflowY: 'auto' }}>
                {updatedStudents.map((s, i) => (
                  <div key={i} style={{
                    display: 'flex', alignItems: 'center', gap: '0.5rem',
                    padding: '0.4rem 0.75rem', borderRadius: '8px',
                    background: 'rgba(16,185,129,0.1)', border: '1px solid rgba(16,185,129,0.2)'
                  }}>
                    <CheckCircle2 size={14} style={{ color: '#10b981', flexShrink: 0 }} />
                    <span style={{ fontSize: '0.85rem', color: '#e2e8f0' }}>{s}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Finish & Download */}
          {mode === 'edit' && updatedCount > 0 && (
            <div className="glass-card">
              <h2 className="section-title"><Download size={18} /> Finish</h2>
              <p style={{ fontSize: '0.82rem', color: '#64748b', marginBottom: '1rem', lineHeight: 1.5 }}>
                All done? Download the updated file. Open it in Excel and your existing formulas (Total, Percentage, Grade) will calculate automatically.
              </p>
              <button className="btn btn-success" onClick={handleDownload} style={{ width: '100%' }}>
                <Download size={18} /> Download Updated Excel
              </button>
            </div>
          )}
        </div>

        {/* ── Main Panel ── */}
        <div className="main-panel">
          {mode === 'upload' && (
            <div className="glass-card" style={{
              display: 'flex', flexDirection: 'column', alignItems: 'center',
              justifyContent: 'center', minHeight: '340px', opacity: 0.5, pointerEvents: 'none'
            }}>
              <FileSpreadsheet size={48} style={{ color: '#334155', marginBottom: '1rem' }} />
              <p style={{ color: '#475569' }}>Upload your Excel template to get started.</p>
            </div>
          )}

          {mode === 'edit' && (
            <div className="glass-card">
              <h2 className="section-title">
                <Sparkles size={20} style={{ color: '#a855f7' }} /> Enter Student Marks
              </h2>

              {/* Subject reference pills */}
              {subjects.length > 0 && (
                <div style={{ marginBottom: '1.25rem' }}>
                  <p style={{ fontSize: '0.78rem', color: '#64748b', marginBottom: '0.5rem' }}>
                    Subjects in this template:
                  </p>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.4rem' }}>
                    {subjects.map((s, i) => (
                      <span key={i} className="badge badge-subject" style={{ fontSize: '0.72rem' }}>{s}</span>
                    ))}
                  </div>
                </div>
              )}

              <form onSubmit={handleProcessPrompt} style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                <textarea
                  className="prompt-textarea"
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  placeholder={`Example: "Eshal Yousaf got 88 in English - Lit + Language, 72 in Mathematics and 90 in General Science."`}
                  disabled={loading || updates.length > 0}
                  rows={3}
                />
                {updates.length === 0 && (
                  <button type="submit" className="btn btn-primary" disabled={loading || !prompt.trim()}>
                    {loading ? <><span className="spinner" /> Parsing...</> : <><Play size={16} /> Parse with AI</>}
                  </button>
                )}
              </form>

              {/* Extracted Updates Preview */}
              {updates.length > 0 && (
                <div style={{ marginTop: '1.75rem' }}>
                  <h3 style={{ fontSize: '0.95rem', color: '#f1f5f9', marginBottom: '0.75rem' }}>
                    AI Extracted — Review before applying:
                  </h3>
                  <div className="table-container" style={{ marginBottom: '1.25rem' }}>
                    <table className="preview-table">
                      <thead>
                        <tr>
                          <th>Student</th>
                          <th>Subject</th>
                          <th>Marks</th>
                          <th style={{ width: '48px' }}></th>
                        </tr>
                      </thead>
                      <tbody>
                        {updates.map((u, idx) => (
                          <tr key={idx}>
                            <td><strong style={{ color: '#e2e8f0' }}>{u.student}</strong></td>
                            <td><span className="badge badge-subject">{u.subject}</span></td>
                            <td style={{ color: '#34d399', fontWeight: 700 }}>{u.marks}</td>
                            <td>
                              <button className="btn-delete-row" onClick={() => handleDeleteUpdate(idx)}>
                                <Trash2 size={14} />
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div style={{ display: 'flex', gap: '0.75rem' }}>
                    <button className="btn btn-success" onClick={handleApplyUpdates} disabled={loading}>
                      {loading ? <><span className="spinner" /> Saving...</> : <><CheckCircle2 size={16} /> Apply to Workbook</>}
                    </button>
                    <button className="btn btn-secondary" onClick={() => { setUpdates([]); setPrompt(''); }} disabled={loading}>
                      Retry
                    </button>
                  </div>
                </div>
              )}

              {/* After applying — add another student */}
              {updates.length === 0 && successMsg && updatedCount > 0 && (
                <div style={{
                  marginTop: '1.5rem', padding: '1rem', borderRadius: '12px',
                  background: 'rgba(99,102,241,0.08)', border: '1px solid rgba(99,102,241,0.2)',
                  display: 'flex', gap: '1rem', alignItems: 'center'
                }}>
                  <CheckCircle2 size={22} style={{ color: '#6366f1', flexShrink: 0 }} />
                  <div>
                    <p style={{ fontSize: '0.9rem', color: '#e2e8f0', fontWeight: 600, margin: 0 }}>
                      Marks saved! Ready for next student.
                    </p>
                    <p style={{ fontSize: '0.8rem', color: '#64748b', margin: '0.2rem 0 0' }}>
                      Type another student's marks above, or download the file when you're done.
                    </p>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
