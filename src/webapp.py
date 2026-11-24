"""
Ship Report Builder — local browser UI.

Runs a tiny localhost HTTP server (Python stdlib only, NO Tk) and opens the
default browser. Native "Browse" dialogs are shown via the OS (osascript on
macOS, PowerShell on Windows), so the user still picks real local paths.

This exists because the system Python's Tk 8.5 renders a black window on modern
macOS. The browser renders reliably everywhere.
"""

from __future__ import annotations

import json
import subprocess
import sys
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from engine import ReportRequest, generate_report
from lookup import lookup_device, format_lookup

IS_MAC = sys.platform == "darwin"
IS_WIN = sys.platform.startswith("win")


# --------------------------------------------------------------------------- #
# Native file/folder pickers
# --------------------------------------------------------------------------- #
def _osascript(script: str) -> str:
    try:
        out = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        return out.stdout.strip()
    except Exception:
        return ""


def _powershell(script: str) -> str:
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", script],
                             capture_output=True, text=True)
        return out.stdout.strip()
    except Exception:
        return ""


def pick(kind: str) -> str:
    """kind = 'file' | 'folder' | 'save'. Returns a path or '' if cancelled."""
    if IS_MAC:
        if kind == "folder":
            return _osascript('POSIX path of (choose folder with prompt "Select the raw-data folder")')
        if kind == "save":
            return _osascript('POSIX path of (choose file name with prompt "Save report as" '
                              'default name "Generated Report.xlsx")')
        return _osascript('POSIX path of (choose file with prompt "Select an Excel file")')
    if IS_WIN:
        if kind == "folder":
            return _powershell(
                "Add-Type -AssemblyName System.Windows.Forms;"
                "$f=New-Object System.Windows.Forms.FolderBrowserDialog;"
                "if($f.ShowDialog() -eq 'OK'){$f.SelectedPath}")
        if kind == "save":
            return _powershell(
                "Add-Type -AssemblyName System.Windows.Forms;"
                "$f=New-Object System.Windows.Forms.SaveFileDialog;"
                "$f.Filter='Excel|*.xlsx';$f.FileName='Generated Report.xlsx';"
                "if($f.ShowDialog() -eq 'OK'){$f.FileName}")
        return _powershell(
            "Add-Type -AssemblyName System.Windows.Forms;"
            "$f=New-Object System.Windows.Forms.OpenFileDialog;"
            "$f.Filter='Excel|*.xlsx;*.xlsm;*.xls';"
            "if($f.ShowDialog() -eq 'OK'){$f.FileName}")
    return ""


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
  .row{display:flex;gap:8px}
  input[type=text]{flex:1;padding:9px 11px;border:1px solid #ccd;border-radius:8px;font-size:14px}
  button{padding:9px 14px;border:0;border-radius:8px;font-size:14px;cursor:pointer}
  .browse{background:#e8eaf0;color:#334}
  .primary{background:#2f6df6;color:#fff;font-weight:600;width:100%;margin-top:22px;padding:12px}
  .ghost{background:#eef;color:#2f6df6}
  hr{border:0;border-top:1px solid #eee;margin:22px 0}
  #log{white-space:pre-wrap;background:#0f1220;color:#cde;border-radius:8px;padding:14px;margin-top:16px;font-family:Menlo,monospace;font-size:12.5px;min-height:40px}
  .two{display:flex;gap:14px}.two>div{flex:1}
  .hint{color:#889;font-size:12px;font-weight:400}
</style></head><body>
<div class="card">
  <h1>Ship Report Builder</h1>
  <p class="sub">Generate a certification report from raw test files — in your sample's exact format.</p>

  <label>Sample report (template) <span class="hint">— the format to copy</span></label>
  <div class="row"><input type="text" id="sample" placeholder="Choose the sample .xlsx"><button class="browse" onclick="pick('file','sample')">Browse…</button></div>

  <div class="two">
    <div><label>PO Number</label><input type="text" id="po" placeholder="e.g. 1218"></div>
    <div><label>Build Code</label><input type="text" id="build" placeholder="e.g. ND"></div>
  </div>

  <label>Raw-data folder <span class="hint">— folder of test files</span></label>
  <div class="row"><input type="text" id="raw" placeholder="Choose the raw-data folder"><button class="browse" onclick="pick('folder','raw')">Browse…</button></div>

  <label>OK-TO-SHIP report <span class="hint">— optional; sets Shipped Status</span></label>
  <div class="row"><input type="text" id="ots" placeholder="Optional"><button class="browse" onclick="pick('file','ots')">Browse…</button></div>

  <label>Save report as</label>
  <div class="row"><input type="text" id="out" placeholder="Choose where to save"><button class="browse" onclick="pick('save','out')">Browse…</button></div>

  <button class="primary" onclick="generate()">Generate Report</button>

  <hr>
  <label>Look up a single device <span class="hint">— by IMEI / DevEUI</span></label>
  <div class="row"><input type="text" id="dev" placeholder="Enter a device ID"><button class="ghost" onclick="lookup()">Look Up</button></div>

  <div id="log">Ready.</div>
</div>
<script>
const $=id=>document.getElementById(id);
const log=t=>{$('log').textContent=t};
async function pick(kind,target){
  log('Opening chooser…');
  const r=await fetch('/pick',{method:'POST',body:JSON.stringify({kind})});
  const j=await r.json();
  if(j.path){$(target).value=j.path; log('Selected: '+j.path);} else {log('(nothing selected)');}
}
async function generate(){
  log('Working… reading raw files and building the report. This can take a few seconds.');
  const body={sample:$('sample').value,po:$('po').value,build:$('build').value,
              raw:$('raw').value,ots:$('ots').value,out:$('out').value};
  const r=await fetch('/generate',{method:'POST',body:JSON.stringify(body)});
  const j=await r.json();
  log(j.message);
}
async function lookup(){
  log('Looking up device…');
  const body={dev:$('dev').value,raw:$('raw').value,po:$('po').value,build:$('build').value};
  const r=await fetch('/lookup',{method:'POST',body:JSON.stringify(body)});
  const j=await r.json();
  log(j.message);
}
</script></body></html>"""


# --------------------------------------------------------------------------- #
# HTTP handler
# --------------------------------------------------------------------------- #
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):        # silence console noise
        pass

    def _send(self, code, body, ctype="application/json"):
        data = body.encode() if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_json(self):
        n = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send(200, PAGE, "text/html; charset=utf-8")
        else:
            self._send(404, "not found", "text/plain")

    def do_POST(self):
        try:
            if self.path == "/pick":
                body = self._read_json()
                self._send(200, json.dumps({"path": pick(body.get("kind", "file"))}))
            elif self.path == "/generate":
                self._send(200, json.dumps(self._generate(self._read_json())))
            elif self.path == "/lookup":
                self._send(200, json.dumps(self._lookup(self._read_json())))
            else:
                self._send(404, json.dumps({"message": "not found"}))
        except Exception:
            self._send(200, json.dumps({"message": "Error:\n" + traceback.format_exc()}))

    # -- actions -------------------------------------------------------- #
    def _generate(self, b):
        msgs = []
        req = ReportRequest(
            sample_path=Path(b.get("sample", "")),
            raw_reports_dir=Path(b.get("raw", "")),
            output_path=Path(b.get("out", "")),
            po_number=b.get("po", "").strip(),
            build_code=b.get("build", "").strip(),
            ok_to_ship_path=Path(b["ots"]) if b.get("ots") else None,
        )
        if not b.get("sample"):
            return {"message": "Please choose a sample report."}
        if not b.get("raw"):
            return {"message": "Please choose the raw-data folder."}
        if not (req.po_number or req.build_code):
            return {"message": "Please enter a PO Number and/or Build Code."}
        if not b.get("out"):
            return {"message": "Please choose where to save the report."}
        generate_report(req, progress=msgs.append)
        return {"message": "\n".join(msgs)}

    def _lookup(self, b):
        if not b.get("dev"):
            return {"message": "Enter a device ID to look up."}
        if not b.get("raw"):
            return {"message": "Choose the raw-data folder first."}
        dl = lookup_device(b["raw"], b["dev"], b.get("po", ""), b.get("build", ""))
        return {"message": format_lookup(dl, b["dev"])}


def main():
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    port = server.server_address[1]
    url = f"http://127.0.0.1:{port}/"
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    print(f"Ship Report Builder running at {url}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
