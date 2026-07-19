import React, { useState, useCallback } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../contexts/AuthContext';
import './AdminUpload.css';

export default function AdminUpload() {
  const { profile } = useAuth();
  const navigate = useNavigate();
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [period, setPeriod] = useState('');
  const [batch, setBatch] = useState('');
  const [status, setStatus] = useState('idle'); // idle, uploading, success, error
  const [errorMsg, setErrorMsg] = useState('');
  const [stats, setStats] = useState(null);

  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDragIn = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      setIsDragging(true);
    }
  }, []);

  const handleDragOut = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.endsWith('.xls') || droppedFile.name.endsWith('.xlsx')) {
        setFile(droppedFile);
      } else {
        setErrorMsg('Please upload an Excel file (.xls or .xlsx)');
      }
      e.dataTransfer.clearData();
    }
  }, []);

  const handleFileSelect = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
  };

  const handleUpload = async () => {
    if (!file || !period || !batch) {
      setErrorMsg('Please provide a file, a period, and a batch ID.');
      return;
    }
    
    setStatus('uploading');
    setErrorMsg('');
    setStats(null);
    
    const formData = new FormData();
    formData.append('file', file);
    formData.append('period', period);
    formData.append('batch', batch);

    try {
      const response = await fetch('/api/upload_visits', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();
      
      if (response.ok) {
        setStatus('success');
        setStats(data.rows_inserted);
      } else {
        setStatus('error');
        setErrorMsg(data.error || 'Failed to upload and process file.');
      }
    } catch (err) {
      setStatus('error');
      setErrorMsg('Network error occurred. Please try again.');
    }
  };

  const autoGenerateBatch = (e) => {
    const val = e.target.value;
    setPeriod(val);
    if (val && !batch) {
      setBatch(val.toLowerCase().replace(/[^a-z0-9]/g, '_'));
    }
  };

  return (
    <div className="upload-page">
      <header className="upload-hdr">
        <div className="upload-hdr-left">
          <button className="abtn abtn-ghost" onClick={() => navigate('/admin')}>← Back to Admin</button>
        </div>
        <div className="upload-hdr-right">
          <span className="upload-who">{profile?.employee_name}</span>
        </div>
      </header>

      <main className="upload-main">
        <div className="upload-container">
          <div className="upload-title-section">
            <h1 className="upload-title">CRM Data Pipeline</h1>
            <p className="upload-subtitle">Upload PulpoPlus exported visits to instantly update the dashboards.</p>
          </div>

          <div className="upload-form">
            <div className="form-row">
              <div className="form-group">
                <label>Period Label</label>
                <input 
                  type="text" 
                  placeholder="e.g. June 2026" 
                  value={period} 
                  onChange={autoGenerateBatch}
                  disabled={status === 'uploading'}
                />
              </div>
              <div className="form-group">
                <label>Batch ID</label>
                <input 
                  type="text" 
                  placeholder="e.g. june_2026" 
                  value={batch} 
                  onChange={(e) => setBatch(e.target.value)}
                  disabled={status === 'uploading'}
                />
              </div>
            </div>

            <div 
              className={`dropzone ${isDragging ? 'dragging' : ''} ${file ? 'has-file' : ''}`}
              onDragEnter={handleDragIn}
              onDragLeave={handleDragOut}
              onDragOver={handleDrag}
              onDrop={handleDrop}
              onClick={() => document.getElementById('file-upload').click()}
            >
              <input 
                id="file-upload" 
                type="file" 
                accept=".xls,.xlsx" 
                style={{ display: 'none' }} 
                onChange={handleFileSelect}
                disabled={status === 'uploading'}
              />
              {file ? (
                <div className="file-preview">
                  <div className="file-icon">📊</div>
                  <div className="file-info">
                    <div className="file-name">{file.name}</div>
                    <div className="file-size">{(file.size / 1024 / 1024).toFixed(2)} MB</div>
                  </div>
                  <button 
                    className="clear-file" 
                    onClick={(e) => { e.stopPropagation(); setFile(null); setStatus('idle'); }}
                    disabled={status === 'uploading'}
                  >
                    ✕
                  </button>
                </div>
              ) : (
                <div className="dropzone-content">
                  <div className="dropzone-icon">☁️</div>
                  <div className="dropzone-text">
                    <span className="bold-text">Click to upload</span> or drag and drop
                  </div>
                  <div className="dropzone-hint">Excel files (.xls, .xlsx) from PulpoPlus Actual Visits Report</div>
                </div>
              )}
            </div>

            {errorMsg && (
              <div className="upload-alert error">
                {errorMsg}
              </div>
            )}

            {status === 'success' && (
              <div className="upload-alert success">
                <div className="success-icon">✓</div>
                <div className="success-content">
                  <strong>Upload and processing complete!</strong>
                  <div className="stats-grid">
                    <div className="stat-item">
                      <span className="stat-val">{stats?.summaries || 0}</span>
                      <span className="stat-label">Summaries</span>
                    </div>
                    <div className="stat-item">
                      <span className="stat-val">{stats?.coaching_days || 0}</span>
                      <span className="stat-label">Coaching</span>
                    </div>
                    <div className="stat-item">
                      <span className="stat-val">{stats?.specialty_classification || 0}</span>
                      <span className="stat-label">Specialties</span>
                    </div>
                    <div className="stat-item">
                      <span className="stat-val">{stats?.product_calls || 0}</span>
                      <span className="stat-label">Products</span>
                    </div>
                  </div>
                </div>
              </div>
            )}

            <button 
              className={`upload-btn ${status === 'uploading' ? 'loading' : ''}`} 
              onClick={handleUpload}
              disabled={!file || !period || !batch || status === 'uploading'}
            >
              {status === 'uploading' ? (
                <>
                  <div className="spinner"></div>
                  Processing & Uploading...
                </>
              ) : (
                'Sync to Supabase'
              )}
            </button>
          </div>
        </div>
      </main>
    </div>
  );
}
