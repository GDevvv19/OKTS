"""
Ship Report Builder — hosted web UI.

Unlike webapp.py (localhost-only, native OS file/folder pickers), this is
meant to be deployed to a real host and reached from any device: inputs are
browser file uploads (the raw-data folder is uploaded as a single .zip —
engine.py already supports reading raw data straight out of a zip), and the
generated report comes back as a normal file download instead of being saved
to a locally-picked path.

Gated by a single shared password (set via the SHIPREPORT_PASSWORD
environment variable on the host — never hardcoded here). If that variable
isn't set, the app refuses to serve anything rather than silently running
open to the internet.
"""

from __future__ import annotations

import io
import os
import secrets
import shutil
import tempfile
from pathlib import Path

from flask import Flask, request, session, redirect, url_for, send_file, Response, jsonify
from werkzeug.exceptions import HTTPException

from engine import ReportRequest, generate_report
from lookup import lookup_device, format_lookup

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024   # 300 MB upload cap

PASSWORD = os.environ.get("SHIPREPORT_PASSWORD")


# Any failure — a too-large upload, a timeout, an unexpected exception — must
# come back as JSON with a readable message, never an HTML error page. The
# frontend's /generate handler previously assumed "not JSON" meant "this is
# the report," so an HTML error page got silently saved as a .xlsx file (the
# exact "Excel cannot open the file" bug this was built to prevent).
@app.errorhandler(HTTPException)
def _json_http_error(e):
    msg = f"Error {e.code}: {e.description}"
    if e.code == 413:
        msg = "Error: that upload is larger than this server allows (300 MB)."
    return jsonify({"message": msg}), e.code


@app.errorhandler(Exception)
def _json_server_error(e):
    return jsonify({"message": f"Error:\n{e}"}), 500


# --------------------------------------------------------------------------- #
# Password gate
# --------------------------------------------------------------------------- #
@app.before_request
def _require_login():
    if request.endpoint in ("login", "static"):
        return None
    if PASSWORD is None:
        return Response(
            "Ship Report Builder is not configured: the host is missing the "
            "SHIPREPORT_PASSWORD environment variable. Set it and restart.",
            status=503)
    if not session.get("ok"):
        return redirect(url_for("login"))
    return None


LOGIN_PAGE = """<!doctype html><html><head><meta charset="utf-8">
<title>Ship Report Builder — Sign in</title>
<style>
  body{{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#f4f5f7;color:#1a1a2e;
       display:flex;align-items:center;justify-content:center;height:100vh;margin:0}}
  .card{{background:#fff;border-radius:12px;padding:30px 34px;box-shadow:0 2px 14px rgba(0,0,0,.08);width:320px}}
  h1{{margin:0 0 18px;font-size:18px}}
  input{{width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #ccd;border-radius:8px;font-size:14px;margin-bottom:14px}}
  button{{width:100%;padding:11px;border:0;border-radius:8px;background:#2f6df6;color:#fff;font-weight:600;font-size:14px;cursor:pointer}}
  .err{{color:#c0392b;font-size:13px;margin:-8px 0 14px}}
</style></head><body>
<div class="card">
  <h1>Ship Report Builder</h1>
  <form method="post">
    {error}
    <input type="password" name="password" placeholder="Password" autofocus>
    <button type="submit">Sign in</button>
  </form>
</div></body></html>"""


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        if PASSWORD is not None and secrets.compare_digest(request.form.get("password", ""), PASSWORD):
            session["ok"] = True
            return redirect(url_for("index"))
        error = '<p class="err">Incorrect password.</p>'
    return LOGIN_PAGE.format(error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# --------------------------------------------------------------------------- #
# HTML page
# --------------------------------------------------------------------------- #
PAGE = """<!doctype html><html><head><meta charset="utf-8">
<title>Ship Report Builder</title>
<style>
  :root{color-scheme:light}
  body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#f4f5f7;color:#1a1a2e;margin:0;padding:28px}
  .card{max-width:820px;margin:0 auto;background:#fff;border-radius:12px;padding:26px 30px;box-shadow:0 2px 14px rgba(0,0,0,.08)}
  h1{margin:0 0 4px;font-size:22px}
  .sub{color:#667;margin:0 0 20px;font-size:13px}
  label{display:block;font-weight:600;font-size:13px;margin:14px 0 4px}
  input[type=text],input[type=file]{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #ccd;border-radius:8px;font-size:14px}
  button{padding:9px 14px;border:0;border-radius:8px;font-size:14px;cursor:pointer}
  .primary{background:#2f6df6;color:#fff;font-weight:600;width:100%;margin-top:22px;padding:12px}
  .ghost{background:#eef;color:#2f6df6}
  hr{border:0;border-top:1px solid #eee;margin:22px 0}
  #log{white-space:pre-wrap;background:#0f1220;color:#cde;border-radius:8px;padding:14px;margin-top:16px;font-family:Menlo,monospace;font-size:12.5px;min-height:40px}
  .two{display:flex;gap:14px}.two>div{flex:1}
  .hint{color:#889;font-size:12px;font-weight:400}
  .top{display:flex;justify-content:space-between;align-items:baseline}
  .top a{color:#889;font-size:12px}
</style></head><body>
<div class="card">
  <div class="top"><h1>Ship Report Builder</h1><a href="/logout">Sign out</a></div>
  <p class="sub">Generate a certification report from raw test files — in your sample's exact format.</p>

  <form id="genForm">
    <label>Sample report (template) <span class="hint">— the format to copy</span></label>
    <input type="file" name="sample" accept=".xlsx,.xlsm,.xls" required>

    <div class="two">
      <div><label>PO Number</label><input type="text" name="po" placeholder="e.g. 1218"></div>
      <div><label>Build Code</label><input type="text" name="build" placeholder="e.g. ND"></div>
    </div>

    <label>Raw-data folder <span class="hint">— upload as a single .zip</span></label>
    <input type="file" name="raw" accept=".zip" required>

    <label>OK-TO-SHIP report <span class="hint">— optional; sets Shipped Status</span></label>
    <input type="file" name="ots" accept=".xlsx,.xlsm,.xls">

    <label>Save report as <span class="hint">— filename only</span></label>
    <input type="text" name="out" placeholder="Generated Report.xlsx">

    <button type="submit" class="primary">Generate Report</button>
  </form>

  <hr>
  <form id="lookupForm">
    <label>Look up a single device <span class="hint">— by IMEI / DevEUI (upload the raw-data .zip again)</span></label>
    <input type="text" name="dev" placeholder="Enter a device ID">
    <input type="file" name="raw" accept=".zip" style="margin-top:8px">
    <button type="submit" class="ghost" style="margin-top:10px">Look Up</button>
  </form>

  <div id="log">Ready.</div>
</div>
<script>
const $=sel=>document.querySelector(sel);
const log=t=>{$('#log').textContent=t};

$('#genForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  log('Working… uploading files and building the report. This can take a while for large raw-data sets.');
  const fd = new FormData(e.target);
  let r;
  try {
    r = await fetch('/generate', {method:'POST', body: fd});
  } catch (err) {
    log('Error: could not reach the server (' + err + ').');
    return;
  }
  // Any non-success response is an error, regardless of its content-type —
  // never assume "not JSON" means "this must be the report file" (that's
  // what let an HTML error page get silently saved as a .xlsx before).
  if (!r.ok) {
    let msg = 'Error ' + r.status + ': the server could not generate the report.';
    try {
      const j = await r.json();
      if (j && j.message) msg = j.message;
    } catch (err) { /* response wasn't JSON; keep the generic message */ }
    log(msg);
    return;
  }
  if (r.headers.get('Content-Type') === 'application/json') {
    const j = await r.json();
    log(j.message);
    return;
  }
  const blob = await r.blob();
  const name = (r.headers.get('Content-Disposition') || '').match(/filename="?([^"]+)"?/);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name ? name[1] : 'Generated Report.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  log('Done — report downloaded.');
});

$('#lookupForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  log('Looking up device…');
  const fd = new FormData(e.target);
  const r = await fetch('/lookup', {method:'POST', body: fd});
  const j = await r.json();
  log(j.message);
});
</script></body></html>"""


@app.route("/")
def index():
    return PAGE


# --------------------------------------------------------------------------- #
# Actions
# --------------------------------------------------------------------------- #
def _save_upload(file_storage, dest_dir: Path) -> Path:
    dest = dest_dir / file_storage.filename
    file_storage.save(dest)
    return dest


@app.route("/generate", methods=["POST"])
def generate():
    sample = request.files.get("sample")
    raw = request.files.get("raw")
    ots = request.files.get("ots")
    po = request.form.get("po", "").strip()
    build = request.form.get("build", "").strip()
    out_name = request.form.get("out", "").strip() or "Generated Report.xlsx"

    if not sample or not sample.filename:
        return {"message": "Please choose a sample report."}
    if not raw or not raw.filename:
        return {"message": "Please choose the raw-data .zip."}
    if not (po or build):
        return {"message": "Please enter a PO Number and/or Build Code."}

    tmp = Path(tempfile.mkdtemp(prefix="shipreport_"))
    try:
        sample_path = _save_upload(sample, tmp)
        raw_dir = tmp / "raw"
        raw_dir.mkdir()
        _save_upload(raw, raw_dir)
        ots_path = _save_upload(ots, tmp) if ots and ots.filename else None

        msgs = []
        req = ReportRequest(
            sample_path=sample_path,
            raw_reports_dir=raw_dir,
            output_path=tmp / out_name,
            po_number=po,
            build_code=build,
            ok_to_ship_path=ots_path,
        )
        out_path = generate_report(req, progress=msgs.append)
        data = out_path.read_bytes()
    except Exception as e:
        return {"message": f"Error:\n{e}"}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=out_path.name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/lookup", methods=["POST"])
def lookup():
    dev = request.form.get("dev", "").strip()
    raw = request.files.get("raw")
    po = request.form.get("po", "").strip()
    build = request.form.get("build", "").strip()

    if not dev:
        return {"message": "Enter a device ID to look up."}
    if not raw or not raw.filename:
        return {"message": "Choose the raw-data .zip first."}

    tmp = Path(tempfile.mkdtemp(prefix="shipreport_lookup_"))
    try:
        _save_upload(raw, tmp)
        dl = lookup_device(tmp, dev, po, build)
        return {"message": format_lookup(dl, dev)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    # Local dev only — a real deploy runs this under gunicorn (see Procfile).
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
