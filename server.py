"""
Open Files Viewer - Currently open file path viewer
Run: python server.py
Browser opens automatically at http://localhost:8765
"""

import http.server
import json
import os
import subprocess
import socketserver
import threading
import webbrowser
from pathlib import Path

PORT = 8765

# PowerShell script to get open Office files via COM automation
PS_OFFICE_SCRIPT = r"""
$results = @()

# PowerPoint
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    foreach($p in $app.Presentations) {
        if ($p.FullName -and $p.FullName -ne '') {
            $results += [PSCustomObject]@{
                app = 'PowerPoint'
                name = $p.Name
                path = $p.FullName
                ext = '.pptx'
            }
        }
    }
} catch {}

# Word
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    foreach($d in $app.Documents) {
        if ($d.FullName -and $d.FullName -ne '') {
            $results += [PSCustomObject]@{
                app = 'Word'
                name = $d.Name
                path = $d.FullName
                ext = '.docx'
            }
        }
    }
} catch {}

# Excel
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    foreach($w in $app.Workbooks) {
        if ($w.FullName -and $w.FullName -ne '') {
            $results += [PSCustomObject]@{
                app = 'Excel'
                name = $w.Name
                path = $w.FullName
                ext = '.xlsx'
            }
        }
    }
} catch {}

if ($results.Count -eq 0) {
    Write-Output '[]'
} elseif ($results.Count -eq 1) {
    Write-Output ('[' + ($results | ConvertTo-Json -Compress) + ']')
} else {
    Write-Output ($results | ConvertTo-Json -Compress)
}
"""

# PowerShell script to get other open files from process command lines
PS_PROCESS_SCRIPT = r"""
$fileExts = '\.(pdf|txt|csv|log|xml|json|yaml|yml|ini|cfg|conf|html|htm|css|js|ts|py|java|cpp|c|h|cs|go|rs|rb|php|sql|md|rtf|odt|ods|odp|hwp|hwpx|psd|ai|svg|png|jpg|jpeg|gif|bmp|tiff|webp|mp4|avi|mkv|mov|wmv|mp3|wav|flac|aac|zip|rar|7z|tar|gz|iso|dwg|dxf|sketch|fig|xd|indd|epub|mobi)["''\s]?$'

$knownOffice = @('POWERPNT.EXE','WINWORD.EXE','EXCEL.EXE')

$results = @()

Get-CimInstance Win32_Process | Where-Object {
    $_.CommandLine -and
    $_.Name -notin $knownOffice -and
    $_.CommandLine -match $fileExts
} | ForEach-Object {
    $cmdLine = $_.CommandLine
    $procName = $_.Name

    # Extract file paths from command line
    $matches2 = [regex]::Matches($cmdLine, '"([^"]+\.[a-zA-Z0-9]+)"')
    foreach ($m in $matches2) {
        $filePath = $m.Groups[1].Value
        if ((Test-Path $filePath -ErrorAction SilentlyContinue) -and !(Test-Path $filePath -PathType Container -ErrorAction SilentlyContinue)) {
            $ext = [System.IO.Path]::GetExtension($filePath)
            $results += [PSCustomObject]@{
                app = $procName -replace '\.exe$',''
                name = [System.IO.Path]::GetFileName($filePath)
                path = $filePath
                ext = $ext.ToLower()
            }
        }
    }

    # Also check unquoted paths
    if ($matches2.Count -eq 0) {
        $parts = $cmdLine -split '\s+'
        foreach ($part in $parts) {
            $part = $part.Trim('"', "'")
            if ($part -match $fileExts -and (Test-Path $part -ErrorAction SilentlyContinue) -and !(Test-Path $part -PathType Container -ErrorAction SilentlyContinue)) {
                $ext = [System.IO.Path]::GetExtension($part)
                $results += [PSCustomObject]@{
                    app = $procName -replace '\.exe$',''
                    name = [System.IO.Path]::GetFileName($part)
                    path = $part
                    ext = $ext.ToLower()
                }
            }
        }
    }
}

if ($results.Count -eq 0) {
    Write-Output '[]'
} elseif ($results.Count -eq 1) {
    Write-Output ('[' + ($results | ConvertTo-Json -Compress) + ']')
} else {
    Write-Output ($results | ConvertTo-Json -Compress)
}
"""


def run_powershell(script: str) -> list:
    """Run a PowerShell script and return parsed JSON result."""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        output = result.stdout.strip()
        if not output or output == "[]":
            return []
        return json.loads(output)
    except (subprocess.TimeoutExpired, json.JSONDecodeError, Exception) as e:
        print(f"PowerShell error: {e}")
        return []


def get_open_files() -> list:
    """Get all currently open files."""
    seen_paths = set()
    files = []

    # Get Office files via COM
    office_files = run_powershell(PS_OFFICE_SCRIPT)
    for f in office_files:
        path = f.get("path", "")
        if path and path not in seen_paths:
            seen_paths.add(path)
            files.append(f)

    # Get other files from process command lines
    process_files = run_powershell(PS_PROCESS_SCRIPT)
    for f in process_files:
        path = f.get("path", "")
        if path and path not in seen_paths:
            seen_paths.add(path)
            files.append(f)

    return files


class RequestHandler(http.server.SimpleHTTPRequestHandler):
    """HTTP request handler with API endpoints."""

    def do_GET(self):
        if self.path == "/api/files":
            self.send_json_response(get_open_files())
        elif self.path == "/" or self.path == "/index.html":
            self.serve_file("index.html")
        else:
            super().do_GET()

    def send_json_response(self, data):
        response = json.dumps(data, ensure_ascii=False)
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(response.encode("utf-8"))

    def serve_file(self, filename):
        filepath = Path(__file__).parent / filename
        if filepath.exists():
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(filepath.read_bytes())
        else:
            self.send_error(404)

    def log_message(self, format, *args):
        pass  # Suppress request logs


def main():
    os.chdir(Path(__file__).parent)

    with socketserver.TCPServer(("", PORT), RequestHandler) as httpd:
        httpd.allow_reuse_address = True
        url = f"http://localhost:{PORT}"
        print(f"Open Files Viewer running at {url}")
        print("Press Ctrl+C to stop.")

        # Open browser after a short delay
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()

        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nStopped.")
            httpd.shutdown()


if __name__ == "__main__":
    main()
