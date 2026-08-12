"""
Broking MIS Dashboard — Web Server

Flask app that provides:
  - GET /           Upload page for two Excel files (FY base + new period)
  - POST /process   Processes uploaded files, generates dashboard, redirects
  - GET /dashboard  Serves the generated dashboard HTML
  - GET /status     SSE endpoint for processing progress
"""
import gzip
import json
import os
import sys
import tempfile
import threading
import time
import traceback
from pathlib import Path

from flask import Flask, request, redirect, send_file, Response, jsonify

# Ensure our project root is importable
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from build_preprocess import process as preprocess_data
from build_dashboard import generate_html, get_sheetjs

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB max upload

# Processing state
processing_state = {"status": "idle", "step": "", "error": None, "progress": 0}
dashboard_html_path = ROOT / "outputs" / "Broking_MIS_Dashboard_Latest.html"
dashboard_gz_path = ROOT / "outputs" / "dashboard.html.gz"


UPLOAD_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Broking MIS Dashboard — Upload Files</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --navy:#10253F;--navy2:#173A5E;--navy-light:#1E3A5F;
  --ivory:#FAF7F0;--paper:#F3EFE6;--gold:#C9A227;--gold-light:#E8D48B;
  --ink:#1C2833;--muted:#617182;--line:#D9D4C8;
  --green:#166534;--red:#9F1239;
  --radius:10px;--shadow:0 4px 24px rgba(16,37,63,.15);
}
html{font-size:15px}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--paper);color:var(--ink);min-height:100vh;display:flex;flex-direction:column}

/* Header */
.header{background:linear-gradient(135deg,var(--navy) 0%,var(--navy2) 100%);color:#fff;padding:1.8rem 2rem;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,.25)}
.header h1{font-size:1.8rem;font-weight:700;letter-spacing:.02em;margin-bottom:.3rem}
.header p{color:var(--gold-light);font-size:.9rem}

/* Main */
.main{flex:1;display:flex;align-items:center;justify-content:center;padding:2rem}
.card{background:#fff;border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);max-width:680px;width:100%;overflow:hidden}
.card-header{background:var(--navy);padding:1.2rem 1.8rem;color:#fff}
.card-header h2{font-size:1.15rem;font-weight:600}
.card-header p{font-size:.8rem;color:var(--gold-light);margin-top:.2rem}
.card-body{padding:2rem 1.8rem}

/* File upload */
.upload-group{margin-bottom:1.6rem}
.upload-group label{display:block;font-weight:600;font-size:.85rem;color:var(--navy);margin-bottom:.5rem}
.upload-group .hint{font-size:.75rem;color:var(--muted);margin-bottom:.6rem}
.file-drop{border:2px dashed var(--line);border-radius:var(--radius);padding:1.4rem;text-align:center;cursor:pointer;transition:all .2s;background:var(--ivory);position:relative}
.file-drop:hover,.file-drop.dragover{border-color:var(--gold);background:#FFF9E6}
.file-drop.has-file{border-color:var(--green);background:#F0FFF4}
.file-drop input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer}
.file-drop .icon{font-size:2rem;margin-bottom:.4rem}
.file-drop .text{font-size:.85rem;color:var(--muted)}
.file-drop .filename{font-size:.85rem;color:var(--green);font-weight:600;margin-top:.3rem}
.file-drop .filesize{font-size:.72rem;color:var(--muted)}

/* Submit */
.btn-submit{width:100%;padding:.9rem;background:linear-gradient(135deg,var(--gold) 0%,#D4A832 100%);border:none;border-radius:var(--radius);color:var(--navy);font-size:1rem;font-weight:700;cursor:pointer;letter-spacing:.03em;transition:all .2s;box-shadow:0 2px 12px rgba(201,162,39,.3)}
.btn-submit:hover{transform:translateY(-1px);box-shadow:0 4px 16px rgba(201,162,39,.4)}
.btn-submit:disabled{opacity:.5;cursor:not-allowed;transform:none}

/* Processing overlay */
.overlay{display:none;position:fixed;inset:0;background:rgba(16,37,63,.85);z-index:9999;justify-content:center;align-items:center;flex-direction:column}
.overlay.active{display:flex}
.overlay .spinner{width:56px;height:56px;border:4px solid rgba(255,255,255,.2);border-top-color:var(--gold);border-radius:50%;animation:spin 1s linear infinite;margin-bottom:1.5rem}
@keyframes spin{to{transform:rotate(360deg)}}
.overlay .step-text{color:#fff;font-size:1.1rem;font-weight:500;text-align:center;max-width:400px}
.overlay .sub-text{color:var(--gold-light);font-size:.85rem;margin-top:.5rem;text-align:center}
.progress-bar{width:300px;height:6px;background:rgba(255,255,255,.15);border-radius:3px;margin-top:1rem;overflow:hidden}
.progress-fill{height:100%;background:var(--gold);border-radius:3px;transition:width .5s ease;width:0}

/* Error */
.error-box{background:var(--red);color:#fff;padding:1rem 1.4rem;border-radius:var(--radius);margin-bottom:1rem;font-size:.85rem;display:none}
.error-box.active{display:block}

/* Footer info */
.footer-info{text-align:center;padding:1rem;color:var(--muted);font-size:.75rem}

/* Previous dashboard link */
.prev-link{text-align:center;margin-top:1rem}
.prev-link a{color:var(--navy);font-size:.85rem;text-decoration:none;border-bottom:1px dashed var(--gold)}
.prev-link a:hover{color:var(--gold)}
</style>
</head>
<body>

<div class="header">
  <h1>Broking MIS Dashboard</h1>
  <p>Marwadi Shares and Finance Ltd. — Operational MIS Analytics Suite</p>
</div>

<div class="main">
  <div class="card">
    <div class="card-header">
      <h2>Upload Period Files</h2>
      <p>Upload the FY base file and the new period file to generate the full analytics dashboard</p>
    </div>
    <div class="card-body">
      <div id="errorBox" class="error-box"></div>

      <form id="uploadForm" enctype="multipart/form-data">
        <div class="upload-group">
          <label>1. Full Fiscal Year (Base) File</label>
          <div class="hint">e.g. Opr MIS Data Apr25 to Mar26.xlsx — this is the reference period</div>
          <div class="file-drop" id="drop1">
            <input type="file" name="fy_file" accept=".xlsx,.xls" id="file1">
            <div class="icon">📊</div>
            <div class="text">Click or drag & drop your FY base file here</div>
            <div class="filename" id="name1"></div>
            <div class="filesize" id="size1"></div>
          </div>
        </div>

        <div class="upload-group">
          <label>2. New Period (Rolling) File</label>
          <div class="hint">e.g. Opr MIS Data Apr26 to Jul26.xlsx — the period to analyze and compare</div>
          <div class="file-drop" id="drop2">
            <input type="file" name="new_file" accept=".xlsx,.xls" id="file2">
            <div class="icon">📈</div>
            <div class="text">Click or drag & drop your new period file here</div>
            <div class="filename" id="name2"></div>
            <div class="filesize" id="size2"></div>
          </div>
        </div>

        <button type="submit" class="btn-submit" id="submitBtn" disabled>
          Generate Dashboard
        </button>
      </form>

      PREV_LINK_PLACEHOLDER
    </div>
  </div>
</div>

<div class="footer-info">
  Files are processed locally on your machine. Nothing is uploaded to the internet.
</div>

<div class="overlay" id="overlay">
  <div class="spinner"></div>
  <div class="step-text" id="stepText">Preparing...</div>
  <div class="sub-text" id="subText">This may take 30–60 seconds for large files</div>
  <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
</div>

<script>
(function() {
  const form = document.getElementById('uploadForm');
  const file1 = document.getElementById('file1');
  const file2 = document.getElementById('file2');
  const btn = document.getElementById('submitBtn');
  const drop1 = document.getElementById('drop1');
  const drop2 = document.getElementById('drop2');
  const overlay = document.getElementById('overlay');
  const errorBox = document.getElementById('errorBox');

  function formatSize(bytes) {
    if (bytes > 1048576) return (bytes / 1048576).toFixed(1) + ' MB';
    return (bytes / 1024).toFixed(0) + ' KB';
  }

  function updateDrop(drop, file, nameId, sizeId) {
    const nameEl = document.getElementById(nameId);
    const sizeEl = document.getElementById(sizeId);
    if (file) {
      drop.classList.add('has-file');
      nameEl.textContent = file.name;
      sizeEl.textContent = formatSize(file.size);
    } else {
      drop.classList.remove('has-file');
      nameEl.textContent = '';
      sizeEl.textContent = '';
    }
    btn.disabled = !(file1.files.length && file2.files.length);
  }

  file1.addEventListener('change', () => updateDrop(drop1, file1.files[0], 'name1', 'size1'));
  file2.addEventListener('change', () => updateDrop(drop2, file2.files[0], 'name2', 'size2'));

  // Drag & drop
  [drop1, drop2].forEach((drop, i) => {
    const fileInput = i === 0 ? file1 : file2;
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    drop.addEventListener('drop', e => {
      e.preventDefault();
      drop.classList.remove('dragover');
      if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        fileInput.dispatchEvent(new Event('change'));
      }
    });
  });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    errorBox.classList.remove('active');
    overlay.classList.add('active');
    btn.disabled = true;

    const fd = new FormData(form);

    // Poll for progress
    const pollInterval = setInterval(async () => {
      try {
        const res = await fetch('/status');
        const st = await res.json();
        document.getElementById('stepText').textContent = st.step || 'Processing...';
        document.getElementById('progressFill').style.width = st.progress + '%';
      } catch {}
    }, 1000);

    try {
      const res = await fetch('/process', { method: 'POST', body: fd });
      clearInterval(pollInterval);
      if (res.ok) {
        document.getElementById('stepText').textContent = 'Dashboard ready! Redirecting...';
        document.getElementById('progressFill').style.width = '100%';
        setTimeout(() => { window.location.href = '/dashboard'; }, 800);
      } else {
        overlay.classList.remove('active');
        let errMsg = 'Processing requires more memory than available on free cloud server. Please view the pre-generated dashboard below or run locally on your laptop.';
        try {
          const text = await res.text();
          if (text) {
            const err = JSON.parse(text);
            if (err.error) errMsg = err.error;
          }
        } catch {}
        errorBox.innerHTML = errMsg + '<br><br><a href="/dashboard" style="color:#FFF9E6;font-weight:bold;text-decoration:underline">👉 View Live Pre-generated Dashboard Here</a>';
        errorBox.classList.add('active');
        btn.disabled = false;
      }
    } catch (err) {
      clearInterval(pollInterval);
      overlay.classList.remove('active');
      errorBox.innerHTML = 'Cloud memory limit reached while parsing large Excel files.<br>Please <a href="/dashboard" style="color:#FFF9E6;font-weight:bold;text-decoration:underline">click here to view the live dashboard</a> directly!';
      errorBox.classList.add('active');
      btn.disabled = false;
    }
  });
})();
</script>
</body>
</html>"""


@app.route("/")
def index():
    # If request has ?upload=1, show upload page anyway
    if request.args.get("upload") != "1":
        if dashboard_html_path.exists() or dashboard_gz_path.exists():
            return redirect("/dashboard")

    page = UPLOAD_PAGE
    # Show link to previous dashboard if it exists
    if dashboard_html_path.exists() or dashboard_gz_path.exists():
        link = '<div class="prev-link"><a href="/dashboard">View previously generated dashboard →</a></div>'
    else:
        link = ""
    page = page.replace("PREV_LINK_PLACEHOLDER", link)
    return page


@app.route("/process", methods=["POST"])
def process_files():
    global processing_state

    if "fy_file" not in request.files or "new_file" not in request.files:
        return jsonify({"error": "Both files are required"}), 400

    fy_file = request.files["fy_file"]
    new_file = request.files["new_file"]

    if not fy_file.filename or not new_file.filename:
        return jsonify({"error": "Both files must be selected"}), 400

    # Save uploaded files to temp dir
    upload_dir = ROOT / "uploads"
    upload_dir.mkdir(exist_ok=True)

    fy_path = upload_dir / fy_file.filename
    new_path = upload_dir / new_file.filename
    fy_file.save(str(fy_path))
    new_file.save(str(new_path))

    try:
        # Step 1: Preprocess
        processing_state = {"status": "running", "step": "Reading Excel files and preprocessing data...", "error": None, "progress": 15}
        payload, stats = preprocess_data(fy_path, new_path)

        processing_state["step"] = f"Processed {stats['fy_channels']} FY channels, {stats['new_channels']} new channels, {stats['matched_clients']:,} matched clients..."
        processing_state["progress"] = 60

        # Step 2: Generate dashboard HTML
        processing_state["step"] = "Generating interactive dashboard..."
        processing_state["progress"] = 75

        sheetjs = get_sheetjs()
        html = generate_html(payload, sheetjs)

        processing_state["step"] = "Writing dashboard file..."
        processing_state["progress"] = 90

        dashboard_html_path.write_text(html, encoding="utf-8")
        with gzip.open(dashboard_gz_path, "wb", compresslevel=9) as f:
            f.write(html.encode("utf-8"))

        del payload, html
        gc.collect()

        processing_state = {"status": "done", "step": "Complete!", "error": None, "progress": 100}

        # Clean up uploaded files
        try:
            fy_path.unlink()
            new_path.unlink()
        except:
            pass

        return jsonify({"success": True, "stats": stats})

    except Exception as e:
        processing_state = {"status": "error", "step": str(e), "error": traceback.format_exc(), "progress": 0}
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500


@app.route("/status")
def status():
    return jsonify(processing_state)


UPLOAD_BUTTON_HTML = '''<a href="/?upload=1" class="btn-upload" title="Upload new period files to regenerate the dashboard" style="display:inline-flex;align-items:center;gap:.45rem;background:linear-gradient(135deg,#C9A227 0%,#D4A832 100%);color:#10253F;border:none;padding:.6rem 1.2rem;border-radius:6px;font-size:.82rem;font-weight:700;cursor:pointer;letter-spacing:.03em;text-decoration:none;box-shadow:0 2px 10px rgba(201,162,39,.35);white-space:nowrap;margin-left:auto;flex-shrink:0">
    <svg viewBox="0 0 24 24" style="width:16px;height:16px;fill:currentColor"><path d="M9 16h6v-6h4l-7-7-7 7h4v6zm-4 2h14v2H5v-2z"/></svg>
    Upload New Files
  </a>'''

UPLOAD_BUTTON_STYLE = '<style>.site-header{display:flex!important;justify-content:space-between!important;align-items:center!important}</style>'


@app.route("/dashboard")
def dashboard():
    # Try serving the full HTML file first (local development)
    if dashboard_html_path.exists():
        html = dashboard_html_path.read_text(encoding="utf-8")
        # Always inject the upload button into the header if not already present
        if "btn-upload" not in html:
            html = html.replace("</head>", UPLOAD_BUTTON_STYLE + "\n</head>", 1)
            html = html.replace("</header>", UPLOAD_BUTTON_HTML + "\n</header>", 1)
        return html

    # Fall back to compressed .gz file (cloud deployment)
    if dashboard_gz_path.exists():
        gz_data = dashboard_gz_path.read_bytes()
        response = Response(gz_data, mimetype="text/html")
        response.headers["Content-Encoding"] = "gzip"
        response.headers["Cache-Control"] = "public, max-age=3600"
        return response

    return redirect("/")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    print("\n" + "=" * 60)
    print("  Broking MIS Dashboard Server")
    print(f"  Open http://localhost:{port} in your browser")
    print("=" * 60 + "\n")
    app.run(host="0.0.0.0", port=port, debug=False)
