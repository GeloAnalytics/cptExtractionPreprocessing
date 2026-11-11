"""table2sheet • Web UI & Local Server
Offline Browser-based interface for table2sheet.
Runs on localhost with zero external dependencies (Python standard library only).
"""

import base64
import http.server
import json
import os
import re
import socket
import socketserver
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
import zipfile
from pathlib import Path
from typing import List, Optional

# Normalize APP_DIR to project root regardless of where web_server.py is located
APP_DIR = Path(__file__).resolve().parent
if APP_DIR.name == "table2sheet_app" and (APP_DIR.parent / "table2sheet.py").exists():
    APP_DIR = APP_DIR.parent

if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

try:
    import table2sheet as t2s
except ImportError:
    from . import table2sheet as t2s

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>table2sheet • PDF & Image Table Extractor</title>
<style>
:root {
  --primary: #4F46E5;
  --primary-hover: #4338CA;
  --primary-light: #EEF2FF;
  --success: #059669;
  --success-hover: #047857;
  --success-bg: #ECFDF5;
  --success-border: #A7F3D0;
  --bg: #F8FAFC;
  --card-bg: #FFFFFF;
  --border: #E2E8F0;
  --text-main: #0F172A;
  --text-muted: #64748B;
  --console-bg: #0F172A;
  --console-text: #E2E8F0;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
html, body {
  width: 100%;
  max-width: 100vw;
  overflow-x: hidden;
}
body {
  font-family: system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  background: var(--bg);
  color: var(--text-main);
  line-height: 1.5;
  padding-bottom: 50px;
}
header {
  background: #FFFFFF;
  border-bottom: 1px solid var(--border);
  padding: 12px 28px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.brand {
  display: flex;
  align-items: center;
  gap: 12px;
}
.logo-badge {
  background: var(--primary);
  color: #FFFFFF;
  font-weight: 800;
  font-size: 15px;
  padding: 6px 12px;
  border-radius: 8px;
  letter-spacing: -0.5px;
}
.brand-text h1 {
  font-size: 15px;
  font-weight: 700;
  color: var(--text-main);
}
.brand-text p {
  font-size: 12px;
  color: var(--text-muted);
}
.header-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}
.status-pill {
  background: var(--success-bg);
  color: var(--success);
  border: 1px solid var(--success-border);
  font-size: 12px;
  font-weight: 700;
  padding: 5px 12px;
  border-radius: 999px;
  display: flex;
  align-items: center;
  gap: 6px;
}
.status-dot {
  width: 8px;
  height: 8px;
  background: var(--success);
  border-radius: 50%;
}
.btn-quit {
  background: none;
  border: 1px solid var(--border);
  color: var(--text-muted);
  font-size: 12px;
  font-weight: 600;
  padding: 5px 12px;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.15s ease;
}
.btn-quit:hover {
  background: #FEE2E2;
  color: #DC2626;
  border-color: #FCA5A5;
}
.container {
  max-width: 1400px;
  width: 100%;
  margin: 24px auto;
  padding: 0 24px;
  display: grid;
  grid-template-columns: minmax(360px, 420px) minmax(0, 1fr);
  gap: 24px;
  box-sizing: border-box;
}
.container > div {
  min-width: 0;
  max-width: 100%;
  width: 100%;
}
@media (max-width: 1024px) {
  .container {
    grid-template-columns: 1fr;
    padding: 0 16px;
  }
}
.card {
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 20px;
  margin-bottom: 20px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}
.card-title {
  font-size: 14px;
  font-weight: 700;
  color: var(--text-main);
  margin-bottom: 14px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.drop-zone {
  border: 2px dashed #CBD5E1;
  border-radius: 10px;
  padding: 26px 16px;
  text-align: center;
  cursor: pointer;
  background: #F8FAFC;
  transition: all 0.2s ease;
}
.drop-zone:hover, .drop-zone.dragover {
  border-color: var(--primary);
  background: var(--primary-light);
}
.drop-icon {
  font-size: 32px;
  color: var(--primary);
  margin-bottom: 6px;
}
.drop-zone p {
  font-size: 13px;
  color: var(--text-muted);
}
.drop-zone p strong {
  color: var(--primary);
}
.drop-actions {
  display: flex;
  gap: 10px;
  justify-content: center;
  margin-top: 12px;
  flex-wrap: wrap;
}
.btn-action-files {
  background: var(--primary);
  color: #FFFFFF;
  border: none;
  border-radius: 8px;
  padding: 8px 18px;
  font-size: 13px;
  font-weight: 700;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: all 0.15s ease;
}
.btn-action-files:hover {
  background: var(--primary-hover);
  transform: translateY(-1px);
}
.btn-action-folder {
  background: #FFFFFF;
  color: var(--primary);
  border: 1.5px solid var(--primary);
  border-radius: 8px;
  padding: 8px 18px;
  font-size: 13px;
  font-weight: 700;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: all 0.15s ease;
}
.btn-action-folder:hover {
  background: var(--primary-light);
  transform: translateY(-1px);
}
.drop-tip {
  font-size: 11px;
  color: #4338CA;
  background: #EEF2FF;
  border: 1px solid #C7D2FE;
  border-radius: 6px;
  padding: 5px 10px;
  margin-top: 10px;
  display: inline-block;
}
.drop-tip kbd {
  background: #FFFFFF;
  border: 1px solid #A5B4FC;
  border-radius: 4px;
  padding: 1px 4px;
  font-family: inherit;
  font-size: 10px;
  font-weight: 700;
}
.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  border: 0;
}
.file-list {
  margin-top: 14px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: 200px;
  overflow-y: auto;
}
.file-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: #F8FAFC;
  border: 1px solid var(--border);
  padding: 8px 12px;
  border-radius: 8px;
  font-size: 13px;
}
.file-item-left {
  display: flex;
  align-items: center;
  gap: 8px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.file-name {
  font-weight: 600;
  color: var(--text-main);
}
.file-size {
  color: var(--text-muted);
  font-size: 11px;
}
.btn-remove {
  background: none;
  border: none;
  color: #EF4444;
  font-size: 16px;
  font-weight: bold;
  cursor: pointer;
  padding: 0 4px;
}
.btn-clear {
  background: none;
  border: none;
  color: var(--text-muted);
  font-size: 12px;
  cursor: pointer;
}
.btn-clear:hover { color: #EF4444; }
.options-grid {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.chk-label {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 13px;
  cursor: pointer;
  user-select: none;
}
.chk-label input[type="checkbox"] {
  width: 17px;
  height: 17px;
  accent-color: var(--primary);
  cursor: pointer;
}
.highlight-opt {
  font-weight: 600;
  color: var(--primary);
}
.format-toggle {
  display: flex;
  gap: 8px;
}
.format-btn {
  flex: 1;
  padding: 8px;
  border: 1px solid var(--border);
  background: #F8FAFC;
  border-radius: 8px;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  text-align: center;
  transition: all 0.15s ease;
}
.format-btn.active {
  background: var(--primary);
  color: #FFFFFF;
  border-color: var(--primary);
}
.adv-toggle {
  font-size: 12px;
  font-weight: 600;
  color: var(--primary);
  cursor: pointer;
  margin-top: 6px;
  display: inline-block;
}
.adv-panel {
  display: none;
  grid-template-columns: 1fr 1fr;
  gap: 12px;
  background: #F8FAFC;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 12px;
  margin-top: 10px;
}
.adv-field label {
  display: block;
  font-size: 11px;
  font-weight: 700;
  color: var(--text-muted);
  margin-bottom: 4px;
}
.adv-field input {
  width: 100%;
  padding: 6px 10px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 12px;
  background: #FFFFFF;
}
.btn-convert {
  width: 100%;
  background: var(--primary);
  color: #FFFFFF;
  border: none;
  border-radius: 10px;
  padding: 14px;
  font-size: 15px;
  font-weight: 700;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 10px;
  box-shadow: 0 2px 4px rgba(79, 70, 229, 0.25);
  transition: all 0.2s ease;
}
.btn-convert:hover:not(:disabled) {
  background: var(--primary-hover);
  transform: translateY(-1px);
}
.btn-convert:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}
.console-box {
  background: var(--console-bg);
  color: var(--console-text);
  border-radius: 10px;
  padding: 14px 16px;
  font-family: Consolas, "Courier New", monospace;
  font-size: 12px;
  height: 240px;
  overflow-y: auto;
  white-space: pre-wrap;
  line-height: 1.6;
}
.console-box .log-success { color: #34D399; font-weight: bold; }
.console-box .log-notice { color: #38BDF8; }
.console-box .log-error { color: #F87171; font-weight: bold; }
.result-card {
  display: none;
  background: #FFFFFF;
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 20px;
  margin-top: 20px;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
  max-width: 100%;
  width: 100%;
  box-sizing: border-box;
  overflow: hidden;
}
.result-banner {
  background: var(--success-bg);
  border: 1.5px solid var(--success-border);
  border-radius: 10px;
  padding: 14px 18px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  margin-bottom: 14px;
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  box-shadow: 0 1px 3px rgba(5, 150, 105, 0.1);
}
.result-banner-info {
  display: flex;
  flex-direction: column;
  gap: 3px;
  min-width: 0;
}
.result-title {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 15px;
  font-weight: 700;
  color: var(--success);
}
.result-summary-text {
  font-size: 13px;
  color: #065F46;
  font-weight: 600;
}
.btn-download {
  background: var(--success);
  color: #FFFFFF;
  text-decoration: none;
  font-size: 14px;
  font-weight: 700;
  padding: 10px 18px;
  border-radius: 8px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  box-shadow: 0 2px 4px rgba(5, 150, 105, 0.25);
  transition: all 0.2s ease;
  flex-shrink: 0;
  white-space: nowrap;
  cursor: pointer;
}
.btn-download:hover {
  background: var(--success-hover);
  transform: translateY(-1px);
  box-shadow: 0 4px 8px rgba(5, 150, 105, 0.35);
}
.btn-cpt-download {
  background: var(--primary);
  color: #FFFFFF;
  text-decoration: none;
  font-size: 14px;
  font-weight: 700;
  padding: 10px 18px;
  border-radius: 8px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  box-shadow: 0 2px 4px rgba(79, 70, 229, 0.25);
  transition: all 0.2s ease;
  white-space: nowrap;
  flex-shrink: 0;
  cursor: pointer;
}
.btn-cpt-download:hover {
  background: var(--primary-hover);
  transform: translateY(-1px);
  box-shadow: 0 4px 8px rgba(79, 70, 229, 0.35);
}
.btn-cpt-download-float {
  background: #3730A3;
  color: #FFFFFF;
  font-size: 13px;
  font-weight: 700;
  text-decoration: none;
  padding: 9px 16px;
  border-radius: 40px;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  box-shadow: 0 2px 6px rgba(55, 48, 163, 0.4);
  transition: all 0.2s ease;
  white-space: nowrap;
}
.btn-cpt-download-float:hover {
  background: #312E81;
  transform: scale(1.03);
}
.preview-tab-btn {
  background: #F1F5F9;
  border: 1px solid var(--border);
  color: var(--text-muted);
  font-size: 12px;
  font-weight: 600;
  padding: 6px 14px;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.15s ease;
}
.preview-tab-btn:hover {
  background: #E2E8F0;
  color: var(--text-main);
}
.preview-tab-btn.active {
  background: var(--primary-light);
  color: var(--primary);
  border-color: #C7D2FE;
  font-weight: 700;
}
.preview-wrap {
  margin-top: 10px;
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow-x: auto;
  overflow-y: auto;
  max-height: 420px;
  width: 100%;
  max-width: 100%;
  display: block;
  box-sizing: border-box;
}
.preview-table {
  width: max-content;
  min-width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}
.preview-table th {
  background: #F1F5F9;
  color: var(--text-main);
  padding: 9px 12px;
  text-align: left;
  font-weight: 700;
  border-bottom: 2px solid var(--border);
  position: sticky;
  top: 0;
  white-space: nowrap;
  z-index: 10;
}
.preview-table td {
  padding: 7px 12px;
  border-bottom: 1px solid #F1F5F9;
  white-space: nowrap;
}
.preview-table tr:nth-child(even) td { background: #FAFAFA; }
.preview-table tr:hover td { background: #EEF2FF; }
.cell-num { text-align: right; font-variant-numeric: tabular-nums; }
.cell-code { font-family: Consolas, monospace; font-weight: 600; color: var(--primary); }
.spinner {
  width: 18px;
  height: 18px;
  border: 3px solid rgba(255,255,255,0.3);
  border-radius: 50%;
  border-top-color: #FFFFFF;
  animation: spin 0.8s linear infinite;
  display: none;
}
@keyframes spin { to { transform: rotate(360deg); } }

/* --- Responsive & Fluid Breakpoints --- */
@media (max-width: 1080px) {
  .container {
    grid-template-columns: 1fr;
    gap: 20px;
    margin: 16px auto;
  }
}

@media (max-width: 768px) {
  header {
    padding: 10px 16px;
    flex-direction: column;
    align-items: stretch;
    gap: 10px;
  }
  .header-actions {
    width: 100%;
    justify-content: space-between;
  }
  .brand-text p {
    display: none;
  }
  .container {
    padding: 0 12px;
    margin: 12px auto;
  }
  .card {
    padding: 14px;
    border-radius: 10px;
  }
  .drop-zone {
    padding: 20px 12px;
  }
  .drop-actions {
    flex-direction: column;
  }
  .btn-action-files, .btn-action-folder {
    width: 100%;
    justify-content: center;
  }
  .adv-panel {
    grid-template-columns: 1fr;
  }
  .result-banner {
    flex-direction: column;
    align-items: stretch;
  }
  .btn-download {
    width: 100%;
    justify-content: center;
  }
}

/* Custom smooth scrollbar for table preview */
.preview-wrap {
  -webkit-overflow-scrolling: touch;
}
.preview-wrap::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}
.preview-wrap::-webkit-scrollbar-track {
  background: #F1F5F9;
  border-radius: 4px;
}
.preview-wrap::-webkit-scrollbar-thumb {
  background: #CBD5E1;
  border-radius: 4px;
}
.preview-wrap::-webkit-scrollbar-thumb:hover {
  background: #94A3B8;
}

/* Floating Fixed Download Pill (Guaranteed Always On-Screen) */
.floating-download-pill {
  position: fixed;
  bottom: 24px;
  right: 28px;
  z-index: 99999;
  background: #064E3B;
  color: #FFFFFF;
  border: 1.5px solid #10B981;
  border-radius: 50px;
  box-shadow: 0 12px 30px -4px rgba(0, 0, 0, 0.4), 0 6px 12px -3px rgba(0, 0, 0, 0.25);
  padding: 8px 12px 8px 18px;
  display: flex;
  align-items: center;
  animation: floatSlideUp 0.35s cubic-bezier(0.16, 1, 0.3, 1);
}
@keyframes floatSlideUp {
  from { transform: translateY(60px); opacity: 0; }
  to { transform: translateY(0); opacity: 1; }
}
.floating-pill-content {
  display: flex;
  align-items: center;
  gap: 14px;
}
.floating-pill-icon {
  background: #10B981;
  color: #FFFFFF;
  width: 28px;
  height: 28px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 800;
  font-size: 14px;
}
.btn-download-float {
  background: #10B981;
  color: #FFFFFF;
  font-size: 13px;
  font-weight: 700;
  text-decoration: none;
  padding: 9px 18px;
  border-radius: 40px;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  box-shadow: 0 2px 6px rgba(16, 185, 129, 0.4);
  transition: all 0.2s ease;
  white-space: nowrap;
}
.btn-download-float:hover {
  background: #059669;
  transform: scale(1.03);
}
</style>
</head>
<body>

<header>
  <div class="brand">
    <div class="logo-badge">table2sheet</div>
    <div class="brand-text">
      <h1>Offline PDF & Image Table to Excel Converter</h1>
      <p>100% private • No cloud APIs • Zero external installation</p>
    </div>
  </div>
  <div class="header-actions">
    <div class="status-pill">
      <div class="status-dot"></div>
      <span>Engine Ready (v2.1 • Web Mode)</span>
    </div>
    <button type="button" class="btn-quit" id="btnQuit" title="Close server process">Quit App</button>
  </div>
</header>

<div class="container">
  <!-- LEFT PANEL -->
  <div>
    <div class="card">
      <div class="card-title">
        <span>Input Documents (<span id="fileCount">0</span> queued)</span>
        <button type="button" class="btn-clear" id="btnClearFiles" style="display:none;">Clear all</button>
      </div>

      <div class="drop-zone" id="dropZone">
        <div class="drop-icon">📄</div>
        <p style="font-size:14px; font-weight:600; color:var(--text-main); margin-bottom:2px;">
          Drag & drop multiple files or a whole folder here
        </p>
        <p style="font-size:11px; color:var(--text-muted); margin-bottom:10px;">
          Supports PDF (.pdf), Excel (.xlsx, .xls), Images (.png, .jpg, .tif)
        </p>

        <div class="drop-actions">
          <button type="button" class="btn-action-files" id="btnBrowseFiles">＋ Select Multiple Files</button>
          <button type="button" class="btn-action-folder" id="btnBrowseFolder">📁 Select Entire Folder</button>
        </div>

        <div class="drop-tip">
          💡 <strong>Tip:</strong> In the file picker, select multiple files by holding <kbd>Ctrl</kbd> or <kbd>Shift</kbd>.
        </div>

        <input type="file" id="fileInput" multiple accept=".pdf,.xlsx,.xls,.csv,.png,.jpg,.jpeg,.tif,.tiff" class="visually-hidden">
        <input type="file" id="folderInput" webkitdirectory directory multiple class="visually-hidden">
      </div>

      <div class="file-list" id="fileList"></div>
    </div>

    <div class="card">
      <div class="card-title">Export Format</div>
      <div class="format-toggle">
        <div class="format-btn active" data-fmt="xlsx">Excel (.xlsx)</div>
        <div class="format-btn" data-fmt="csv">CSV (.csv)</div>
        <div class="format-btn" data-fmt="both">Both (.xlsx + .csv)</div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Extraction Options</div>
      <div class="options-grid">
        <label class="chk-label">
          <input type="checkbox" id="chkNumbers" checked>
          <span>Numbers as real numbers (keeps codes text)</span>
        </label>
        <label class="chk-label">
          <input type="checkbox" id="chkMerge" checked>
          <span>Merge into single sheet</span>
        </label>
        <label class="chk-label">
          <input type="checkbox" id="chkCpt" checked>
          <span class="highlight-opt">★ CPT Mode (Clean Rates & Deduplicate)</span>
        </label>
        <label class="chk-label">
          <input type="checkbox" id="chkOpdOnly" checked>
          <span>OPD rates only (hide ward/room/inpatient columns)</span>
        </label>
        <label class="chk-label">
          <input type="checkbox" id="chkPrecise" checked>
          <span>Precise OCR (Bordered scan multi-tier mode)</span>
        </label>
      </div>

      <span class="adv-toggle" id="btnToggleAdv">▸ Advanced Options</span>
      <div class="adv-panel" id="advPanel">
        <div class="adv-field">
          <label>OCR Language</label>
          <input type="text" id="txtLang" value="eng" placeholder="e.g. eng, fil">
        </div>
        <div class="adv-field">
          <label>Pages (blank = all)</label>
          <input type="text" id="txtPages" placeholder="e.g. 1-3,7">
        </div>
        <div class="adv-field">
          <label>CPT Rate Col Keyword</label>
          <input type="text" id="txtRateCol" placeholder="(blank = all rates)">
        </div>
        <div class="adv-field">
          <label>Excel Sheet Name</label>
          <input type="text" id="txtSheet" placeholder="(auto)">
        </div>
      </div>
    </div>

    <button type="button" class="btn-convert" id="btnConvert">
      <span class="spinner" id="btnSpinner"></span>
      <span id="btnText">⚡ Convert to Spreadsheet</span>
    </button>

    <!-- Always-Visible Quick Download Box on Left Panel -->
    <div id="quickDownloadBox" style="display:none; margin-top:14px;">
      <a href="#" class="btn-cpt-download" id="btnQuickCptDownload" download style="width:100%; justify-content:center; padding:12px; font-size:14px; margin-bottom:8px;">
        <span>🎯</span>
        <span id="quickCptText">Download Tool Ready CSV (.csv)</span>
      </a>
      <a href="#" class="btn-download" id="btnQuickDownload" download style="width:100%; justify-content:center; padding:12px; font-size:14px;">
        <span>⬇</span>
        <span id="quickDownloadText">Download Full Extraction (.xlsx)</span>
      </a>
    </div>
  </div>

  <!-- RIGHT PANEL -->
  <div>
    <div class="card" style="margin-bottom:0;">
      <div class="card-title">
        <span>Activity Console</span>
        <button type="button" class="btn-clear" id="btnClearConsole">Clear</button>
      </div>
      <div class="console-box" id="consoleBox">table2sheet • v2.1 (Browser Mode)
Ready. Drag and drop PDF or image files on the left and click Convert.
</div>
    </div>

    <!-- RESULT & PREVIEW CARD -->
    <div class="result-card" id="resultCard">
      <div class="result-banner">
        <div class="result-banner-info">
          <div class="result-title">
            <span>✓</span>
            <span>Conversion Complete!</span>
          </div>
          <div id="resultSummary" class="result-summary-text">Ready for download</div>
        </div>
        <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
          <a href="#" class="btn-cpt-download" id="btnCptDownload" download title="Strictly formatted for the Maxicare CPT Mapping Tool">
            <span>🎯</span>
            <span id="btnCptText">Download Tool Ready CSV</span>
          </a>
          <a href="#" class="btn-download" id="btnDownload" download title="Full master extraction with all rate columns and duplicates">
            <span>⬇</span>
            <span id="btnDownloadText">Download Full Extraction</span>
          </a>
        </div>
      </div>

      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
        <div style="display:flex; gap:8px;">
          <button type="button" class="preview-tab-btn active" id="tabMasterPreview">📋 Full Extraction Preview</button>
          <button type="button" class="preview-tab-btn" id="tabCptPreview">🎯 Tool Ready (CPT Format) Preview</button>
        </div>
        <div style="font-size:12px; font-weight:600; color:var(--text-muted);" id="previewSubtitle">
          Table Preview (Top 20 rows):
        </div>
      </div>

      <div class="preview-wrap">
        <table class="preview-table" id="previewTable"></table>
      </div>
    </div>
  </div>
</div>

<!-- FLOATING FIXED DOWNLOAD PILL (GUARANTEED ALWAYS ON-SCREEN) -->
<div id="floatingDownloadPill" class="floating-download-pill" style="display:none;">
  <div class="floating-pill-content">
    <div class="floating-pill-icon">✓</div>
    <div class="floating-pill-text">
      <div style="font-weight:700; font-size:13px;" id="floatingPillTitle">Extraction Complete!</div>
      <div style="font-size:11px; opacity:0.9;" id="floatingPillSub">Ready for CPT Mapping Tool</div>
    </div>
    <div style="display:flex; gap:8px; align-items:center;">
      <a href="#" class="btn-cpt-download-float" id="btnFloatingCptDownload" download>
        <span>🎯</span>
        <span>Tool Ready CSV</span>
      </a>
      <a href="#" class="btn-download-float" id="btnFloatingDownload" download>
        <span>⬇</span>
        <span id="floatingDownloadBtnText">Full Extraction</span>
      </a>
    </div>
  </div>
</div>

<script>
let queuedFiles = [];
let exportFormat = "xlsx";

const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const fileList = document.getElementById("fileList");
const btnClearFiles = document.getElementById("btnClearFiles");
const btnConvert = document.getElementById("btnConvert");
const btnSpinner = document.getElementById("btnSpinner");
const btnText = document.getElementById("btnText");
const consoleBox = document.getElementById("consoleBox");
const btnClearConsole = document.getElementById("btnClearConsole");
const resultCard = document.getElementById("resultCard");
const resultSummary = document.getElementById("resultSummary");
const btnDownload = document.getElementById("btnDownload");
const btnCptDownload = document.getElementById("btnCptDownload");
const btnQuickCptDownload = document.getElementById("btnQuickCptDownload");
const btnFloatingCptDownload = document.getElementById("btnFloatingCptDownload");
const tabMasterPreview = document.getElementById("tabMasterPreview");
const tabCptPreview = document.getElementById("tabCptPreview");
const previewTable = document.getElementById("previewTable");
const previewSubtitle = document.getElementById("previewSubtitle");
const btnQuit = document.getElementById("btnQuit");
const btnBrowseFiles = document.getElementById("btnBrowseFiles");
const btnBrowseFolder = document.getElementById("btnBrowseFolder");
const folderInput = document.getElementById("folderInput");
const fileCount = document.getElementById("fileCount");

let currentMasterPreview = [];
let currentCptPreview = [];
let currentTotalRows = 0;

const ALLOWED_EXTS = [".pdf", ".xlsx", ".xls", ".csv", ".png", ".jpg", ".jpeg", ".tif", ".tiff"];

// Format selector
document.querySelectorAll(".format-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".format-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    exportFormat = btn.dataset.fmt;
  });
});

// Advanced panel toggle
const btnToggleAdv = document.getElementById("btnToggleAdv");
const advPanel = document.getElementById("advPanel");
btnToggleAdv.addEventListener("click", () => {
  const isHidden = advPanel.style.display !== "grid";
  advPanel.style.display = isHidden ? "grid" : "none";
  btnToggleAdv.textContent = isHidden ? "▾ Hide Advanced Options" : "▸ Advanced Options";
});

// Explicit buttons
btnBrowseFiles.addEventListener("click", (e) => {
  e.stopPropagation();
  fileInput.click();
});
btnBrowseFolder.addEventListener("click", (e) => {
  e.stopPropagation();
  folderInput.click();
});

// Drop zone clicks (outside buttons)
dropZone.addEventListener("click", (e) => {
  if (e.target.closest("button")) return;
  fileInput.click();
});

fileInput.addEventListener("change", (e) => {
  if (e.target.files && e.target.files.length > 0) {
    addFiles(Array.from(e.target.files));
    fileInput.value = "";
  }
});

folderInput.addEventListener("change", (e) => {
  if (e.target.files && e.target.files.length > 0) {
    addFiles(Array.from(e.target.files));
    folderInput.value = "";
  }
});

// Prevent default on all drag events everywhere
["dragenter", "dragover", "dragleave", "drop"].forEach(name => {
  window.addEventListener(name, (e) => e.preventDefault(), false);
  document.addEventListener(name, (e) => e.preventDefault(), false);
});

window.addEventListener("dragenter", () => dropZone.classList.add("dragover"));
window.addEventListener("dragover", () => dropZone.classList.add("dragover"));
window.addEventListener("dragleave", (e) => {
  if (e.clientX === 0 && e.clientY === 0) {
    dropZone.classList.remove("dragover");
  }
});

window.addEventListener("drop", async (e) => {
  dropZone.classList.remove("dragover");
  const files = await extractDroppedFiles(e.dataTransfer);
  if (files.length > 0) {
    addFiles(files);
  }
});

dropZone.addEventListener("drop", async (e) => {
  e.stopPropagation();
  dropZone.classList.remove("dragover");
  const files = await extractDroppedFiles(e.dataTransfer);
  if (files.length > 0) {
    addFiles(files);
  }
});

// Recursively extract all files even if folders are dropped
async function extractDroppedFiles(dataTransfer) {
  if (!dataTransfer) return [];
  const files = [];

  if (dataTransfer.items && dataTransfer.items.length > 0) {
    const items = Array.from(dataTransfer.items);
    for (const item of items) {
      if (item.kind === "file") {
        const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
        if (entry) {
          if (entry.isFile) {
            const f = item.getAsFile();
            if (f) files.push(f);
          } else if (entry.isDirectory) {
            const dirFiles = await readDirectory(entry);
            files.push(...dirFiles);
          }
        } else {
          const f = item.getAsFile();
          if (f) files.push(f);
        }
      }
    }
  } else if (dataTransfer.files) {
    files.push(...Array.from(dataTransfer.files));
  }
  return files;
}

async function readDirectory(dirEntry) {
  const dirReader = dirEntry.createReader();
  const entries = await new Promise((resolve) => {
    dirReader.readEntries(resolve, () => resolve([]));
  });
  const files = [];
  for (const entry of entries) {
    if (entry.isFile) {
      const f = await new Promise((resolve) => entry.file(resolve, () => resolve(null)));
      if (f) files.push(f);
    } else if (entry.isDirectory) {
      const sub = await readDirectory(entry);
      files.push(...sub);
    }
  }
  return files;
}

function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024, dm = 1;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
}

function addFiles(files) {
  let countAdded = 0;
  for (const f of files) {
    const ext = "." + f.name.split(".").pop().toLowerCase();
    if (!ALLOWED_EXTS.includes(ext)) continue;
    if (!queuedFiles.some(x => x.name === f.name && x.size === f.size)) {
      queuedFiles.push(f);
      countAdded++;
    }
  }
  renderFileList();
  if (countAdded > 0) {
    log(`Added ${countAdded} document(s) to queue. Total: ${queuedFiles.length} file(s).`, "log-notice");
  }
}

function renderFileList() {
  fileList.innerHTML = "";
  fileCount.textContent = queuedFiles.length;
  btnClearFiles.style.display = queuedFiles.length > 0 ? "block" : "none";
  queuedFiles.forEach((f, idx) => {
    const item = document.createElement("div");
    item.className = "file-item";
    item.innerHTML = `
      <div class="file-item-left">
        <span style="font-size:15px;">📄</span>
        <span class="file-name" title="${f.name}">${f.name}</span>
        <span class="file-size">(${formatBytes(f.size)})</span>
      </div>
      <button type="button" class="btn-remove" data-idx="${idx}">×</button>
    `;
    item.querySelector(".btn-remove").addEventListener("click", (e) => {
      e.stopPropagation();
      queuedFiles.splice(idx, 1);
      renderFileList();
    });
    fileList.appendChild(item);
  });
}

btnClearFiles.addEventListener("click", () => {
  queuedFiles = [];
  renderFileList();
});

btnClearConsole.addEventListener("click", () => {
  consoleBox.textContent = "";
});

function log(msg, cls) {
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  consoleBox.appendChild(line);
  consoleBox.scrollTop = consoleBox.scrollHeight;
}

// Convert files to Base64
async function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const b64 = reader.result.split(",")[1];
      resolve(b64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Convert action
btnConvert.addEventListener("click", async () => {
  if (queuedFiles.length === 0) {
    alert("Please add at least one PDF or image file first.");
    return;
  }

  btnConvert.disabled = true;
  btnSpinner.style.display = "inline-block";
  btnText.textContent = "Extracting & Converting...";
  resultCard.style.display = "none";
  document.getElementById("quickDownloadBox").style.display = "none";
  document.getElementById("floatingDownloadPill").style.display = "none";

  log(`Reading ${queuedFiles.map(f => f.name).join(", ")} ...`, "log-notice");

  // Status timer to show active progress
  const progressMessages = [
    "Extracting digital vector tables & cell structures...",
    "Scanning bordered tables with OCR engine...",
    "Harmonizing CPT headers and room rate columns...",
    "Filtering metadata rows and auto-extracting codes...",
    "Converting currency values and removing duplicates...",
  ];
  let pIdx = 0;
  const timer = setInterval(() => {
    if (pIdx < progressMessages.length) {
      log(progressMessages[pIdx++], "log-notice");
    }
  }, 3500);

  try {
    const encodedFiles = [];
    for (const f of queuedFiles) {
      const b64 = await fileToBase64(f);
      encodedFiles.push({ name: f.name, data: b64 });
    }

    const payload = {
      files: encodedFiles,
      format: exportFormat,
      numbers: document.getElementById("chkNumbers").checked,
      merge: document.getElementById("chkMerge").checked,
      cpt: document.getElementById("chkCpt").checked,
      opd_only: document.getElementById("chkOpdOnly").checked,
      precise: document.getElementById("chkPrecise").checked,
      lang: document.getElementById("txtLang").value.trim() || "eng",
      pages: document.getElementById("txtPages").value.trim() || null,
      rate_col: document.getElementById("txtRateCol").value.trim() || null,
      sheet: document.getElementById("txtSheet").value.trim() || null,
    };

    const res = await fetch("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    clearInterval(timer);

    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || "Conversion failed.");
    }

    if (data.logs) {
      data.logs.forEach(l => log(l, "log-notice"));
    }
    log(`✓ Completed successfully: ${data.rows} data rows, ${data.cols} columns extracted.`, "log-success");

    currentMasterPreview = data.preview || [];
    currentCptPreview = data.cpt_preview || [];
    currentTotalRows = data.rows || 0;

    resultSummary.textContent = `Extracted ${data.rows} data rows, ${data.cols} columns`;

    // Master Spreadsheet Download
    btnDownload.href = data.download_url;
    btnDownload.download = data.filename;

    const btnQuickDownload = document.getElementById("btnQuickDownload");
    const quickDownloadBox = document.getElementById("quickDownloadBox");
    const quickDownloadText = document.getElementById("quickDownloadText");
    btnQuickDownload.href = data.download_url;
    btnQuickDownload.download = data.filename;
    quickDownloadText.textContent = `Download ${data.filename}`;

    const floatingPill = document.getElementById("floatingDownloadPill");
    const btnFloatingDownload = document.getElementById("btnFloatingDownload");
    const floatingPillSub = document.getElementById("floatingPillSub");
    btnFloatingDownload.href = data.download_url;
    btnFloatingDownload.download = data.filename;
    floatingPillSub.textContent = `${data.rows} data rows, ${data.cols} columns`;

    // Maxicare CPT Mapping CSV Download
    if (data.cpt_download_url) {
      btnCptDownload.href = data.cpt_download_url;
      btnCptDownload.download = data.cpt_filename;
      btnCptDownload.style.display = "inline-flex";

      btnQuickCptDownload.href = data.cpt_download_url;
      btnQuickCptDownload.download = data.cpt_filename;
      btnQuickCptDownload.style.display = "inline-flex";

      btnFloatingCptDownload.href = data.cpt_download_url;
      btnFloatingCptDownload.download = data.cpt_filename;
      btnFloatingCptDownload.style.display = "inline-flex";

      tabCptPreview.style.display = "inline-block";
    } else {
      btnCptDownload.style.display = "none";
      btnQuickCptDownload.style.display = "none";
      btnFloatingCptDownload.style.display = "none";
      tabCptPreview.style.display = "none";
    }

    quickDownloadBox.style.display = "block";
    floatingPill.style.display = "flex";

    tabMasterPreview.classList.add("active");
    tabCptPreview.classList.remove("active");
    renderPreview(currentMasterPreview, currentTotalRows, "Master Table Preview");

    resultCard.style.display = "block";
    resultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });

  } catch (err) {
    clearInterval(timer);
    log(`✗ ERROR: ${err.message}`, "log-error");
    alert(`Conversion error: ${err.message}`);
  } finally {
    btnConvert.disabled = false;
    btnSpinner.style.display = "none";
    btnText.textContent = "⚡ Convert to Spreadsheet";
  }
});

function renderPreview(table, totalRows, titlePrefix) {
  previewTable.innerHTML = "";
  if (!table || table.length === 0) return;

  const prefix = titlePrefix || "Table Preview";
  previewSubtitle.textContent = `${prefix} (Showing top ${table.length - 1} rows of ${totalRows} total rows):`;

  const thead = document.createElement("thead");
  const trH = document.createElement("tr");
  table[0].forEach(col => {
    const th = document.createElement("th");
    th.textContent = col || "";
    trH.appendChild(th);
  });
  thead.appendChild(trH);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  table.slice(1).forEach(row => {
    const tr = document.createElement("tr");
    row.forEach((cell, ci) => {
      const td = document.createElement("td");
      if (typeof cell === "number") {
        td.className = "cell-num";
        td.textContent = cell.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      } else {
        if (ci === 0 && cell) td.className = "cell-code";
        td.textContent = cell !== null && cell !== undefined ? cell : "";
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  previewTable.appendChild(tbody);
}

// Preview Tab Toggle Handlers
tabMasterPreview.addEventListener("click", () => {
  tabMasterPreview.classList.add("active");
  tabCptPreview.classList.remove("active");
  renderPreview(currentMasterPreview, currentTotalRows, "Master Table Preview");
});

tabCptPreview.addEventListener("click", () => {
  tabCptPreview.classList.add("active");
  tabMasterPreview.classList.remove("active");
  renderPreview(currentCptPreview, currentTotalRows, "CPT Mapping Format Preview (Maxicare Template)");
});

// Quit app action
btnQuit.addEventListener("click", async () => {
  if (confirm("Do you want to stop the table2sheet local server?")) {
    try {
      await fetch("/api/shutdown", { method: "POST" });
    } catch (_) {}
    document.body.innerHTML = `
      <div style="max-width:500px; margin:80px auto; text-align:center; font-family:sans-serif;">
        <h2 style="color:#0F172A; margin-bottom:10px;">table2sheet Server Stopped</h2>
        <p style="color:#64748B;">You can safely close this browser tab.</p>
      </div>
    `;
  }
});
</script>
</body>
</html>
"""


class Table2SheetHandler(http.server.BaseHTTPRequestHandler):
    """Local HTTP Request Handler for table2sheet web interface."""

    def log_message(self, format, *args):
        pass

    def send_cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors_headers()
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            content = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(content)
            return

        if path == "/api/ping":
            data = json.dumps({"status": "ok"}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(data)
            return

        if path.startswith("/api/download/"):
            fname = urllib.parse.unquote(path[len("/api/download/"):])
            download_dir = APP_DIR / "scratch" / "downloads"
            fpath = download_dir / fname
            if fpath.exists() and fpath.is_file():
                self.send_response(200)
                suffix = fpath.suffix.lower()
                if suffix == ".xlsx":
                    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                elif suffix == ".zip":
                    content_type = "application/zip"
                else:
                    content_type = "text/csv; charset=utf-8"
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
                self.send_header("Content-Length", str(fpath.stat().st_size))
                self.send_cors_headers()
                self.end_headers()
                with open(fpath, "rb") as f:
                    while chunk := f.read(65536):
                        self.wfile.write(chunk)
                return
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"File not found")
                return

        self.send_response(404)
        self.end_headers()
        self.wfile.write(b"Not Found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/shutdown":
            self.send_json_response({"status": "shutdown"})
            def _stop():
                time.sleep(0.5)
                os._exit(0)
            threading.Thread(target=_stop, daemon=True).start()
            return

        if parsed.path == "/api/convert":
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            try:
                req = json.loads(body.decode("utf-8"))
            except Exception as e:
                self.send_json_error(f"Invalid JSON request: {e}", 400)
                return

            try:
                result = self.handle_convert(req)
                self.send_json_response(result)
            except Exception as ex:
                import traceback
                traceback.print_exc()
                self.send_json_error(str(ex), 500)
            return

        self.send_response(404)
        self.end_headers()
        self.wfile.write(b"Not Found")

    def handle_convert(self, req: dict) -> dict:
        files_data = req.get("files", [])
        if not files_data:
            raise ValueError("No files provided.")

        fmt = req.get("format", "xlsx")
        upload_dir = APP_DIR / "scratch" / "uploads"
        download_dir = APP_DIR / "scratch" / "downloads"
        upload_dir.mkdir(parents=True, exist_ok=True)
        download_dir.mkdir(parents=True, exist_ok=True)

        saved_paths: List[Path] = []
        for f in files_data:
            name = Path(f["name"]).name
            dest = upload_dir / name
            content = base64.b64decode(f["data"])
            dest.write_bytes(content)
            saved_paths.append(dest)

        logs: List[str] = []
        def web_log(msg: str):
            logs.append(msg)

        opts = t2s.Options(
            numbers=bool(req.get("numbers", True)),
            cpt=bool(req.get("cpt", True)),
            opd_only=bool(req.get("opd_only", True)),
            precise=bool(req.get("precise", True)),
            lang=str(req.get("lang") or "eng").strip(),
            pages=req.get("pages") or None,
            rate_col=req.get("rate_col") or None,
            sheet=req.get("sheet") or None,
            log=web_log,
        )

        out_stem = saved_paths[0].stem if len(saved_paths) == 1 else f"{saved_paths[0].stem}_merged"
        target_dest = download_dir / f"{out_stem}.{fmt}"

        effective_merge = bool(req.get("merge", True)) or (opts.cpt and len(saved_paths) > 1)
        written = t2s.run(
            [str(p) for p in saved_paths],
            fmt=fmt,
            output=str(target_dest),
            merge=effective_merge,
            opts=opts,
        )

        if not written:
            raise RuntimeError("Extraction completed but no file was written.")

        # Identify CPT mapping CSV and primary spreadsheet files
        cpt_file = next((w for w in written if w.name.endswith("_cpt_mapping.csv")), None)
        tool_ready_file = next((w for w in written if w.name.endswith("_tool_ready.csv")), None)
        sheet_files = [w for w in written if not w.name.endswith("_cpt_mapping.csv") and not w.name.endswith("_tool_ready.csv")]
        download_file = sheet_files[0] if sheet_files else written[0]

        if len(sheet_files) > 1:
            zip_dest = download_dir / f"{out_stem}_all_files.zip"
            with zipfile.ZipFile(zip_dest, "w", zipfile.ZIP_DEFLATED) as zf:
                for w in written:
                    zf.write(w, arcname=w.name)
            download_file = zip_dest

        # Generate preview from the primary spreadsheet
        preview: List[List[str]] = []
        primary_xlsx = next((w for w in sheet_files if w.suffix.lower() == ".xlsx"), None)
        try:
            if primary_xlsx:
                import openpyxl
                wb = openpyxl.load_workbook(primary_xlsx, data_only=True)
                ws = wb.active
                for r in range(1, min(22, ws.max_row + 1)):
                    row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                    preview.append(row_vals)
            elif sheet_files:
                import csv
                with open(sheet_files[0], newline="", encoding="utf-8") as cf:
                    reader = csv.reader(cf)
                    for i, r in enumerate(reader):
                        if i >= 21:
                            break
                        preview.append(r)
        except Exception as e:
            web_log(f"Notice generating preview: {e}")

        # Generate preview for the CPT mapping CSV
        cpt_preview: List[List[str]] = []
        if cpt_file and cpt_file.exists():
            try:
                import csv
                with open(cpt_file, newline="", encoding="utf-8-sig") as cf:
                    reader = csv.reader(cf)
                    for i, r in enumerate(reader):
                        if i >= 21:
                            break
                        cpt_preview.append(r)
            except Exception as e:
                web_log(f"Notice generating CPT preview: {e}")

        # Count total data rows across generated spreadsheets
        total_data_rows = 0
        try:
            if primary_xlsx:
                import openpyxl
                wb = openpyxl.load_workbook(primary_xlsx, data_only=True)
                ws = wb.active
                total_data_rows = max(0, ws.max_row - 1)
            elif cpt_file:
                with open(cpt_file, newline="", encoding="utf-8-sig") as cf:
                    total_data_rows = max(0, sum(1 for _ in cf) - 1)
            else:
                total_data_rows = max(0, len(preview) - 1)
        except Exception:
            total_data_rows = max(0, len(preview) - 1)

        cols_count = len(preview[0]) if preview else (len(cpt_preview[0]) if cpt_preview else 0)

        return {
            "success": True,
            "filename": download_file.name,
            "download_url": f"/api/download/{urllib.parse.quote(download_file.name)}",
            "cpt_filename": cpt_file.name if cpt_file else None,
            "cpt_download_url": f"/api/download/{urllib.parse.quote(cpt_file.name)}" if cpt_file else None,
            "rows": total_data_rows,
            "cols": cols_count,
            "logs": logs,
            "preview": preview,
            "cpt_preview": cpt_preview,
        }

    def send_json_response(self, data: dict, code: int = 200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_cors_headers()
        self.end_headers()
        self.wfile.write(body)

    def send_json_error(self, message: str, code: int = 500):
        self.send_json_response({"success": False, "error": message}, code=code)


def is_server_already_running(port: int = 8765) -> bool:
    """Check if table2sheet server is already responding on this port."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=0.8) as resp:
            return resp.status == 200
    except Exception:
        return False


def find_free_port(start_port: int = 8765, max_attempts: int = 50) -> int:
    for port in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("No free port available in range.")


def start_server(port: int = 8765, open_browser: bool = True):
    """Start local table2sheet web server and open in default browser."""
    # If already running on standard port, send shutdown signal so fresh code can start
    if is_server_already_running(port):
        print(f"table2sheet is already running on port {port}. Refreshing server instance...")
        try:
            req = urllib.request.Request(
                f"http://127.0.0.1:{port}/api/shutdown",
                data=b"{}",
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=1.0):
                pass
            time.sleep(0.8)
        except Exception:
            pass

    free_port = port if not is_server_already_running(port) else find_free_port(port)
    url = f"http://localhost:{free_port}"

    server = socketserver.TCPServer(("127.0.0.1", free_port), Table2SheetHandler)
    server.allow_reuse_address = True

    print("=" * 65)
    print("  table2sheet • Web UI Server (v2.1)")
    print(f"  Running at: {url}")
    print("  Opening in your default browser...")
    print("  Close this console window or click 'Quit App' to stop.")
    print("=" * 65)

    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\\nShutting down server...")
    finally:
        server.server_close()


if __name__ == "__main__":
    start_server()
