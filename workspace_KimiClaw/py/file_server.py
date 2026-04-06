#!/usr/bin/env python3
"""
Simple HTTP server with file upload support
Usage: python3 upload_server.py [port]
"""
import http.server
import socketserver
import os
import cgi
import sys
from pathlib import Path

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
UPLOAD_DIR = Path("/root/.openclaw/workspace")

class UploadHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        """Handle GET requests - list files or serve file"""
        path = UPLOAD_DIR / self.path.lstrip('/')
        
        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            
            html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>OpenClaw Workspace File Manager</title>
    <meta charset="utf-8">
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 1200px; margin: 40px auto; padding: 0 20px; background: #f5f5f5; }}
        h1 {{ color: #333; border-bottom: 2px solid #0066cc; padding-bottom: 10px; }}
        .upload-box {{ background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 30px; }}
        .file-list {{ background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #eee; }}
        th {{ background: #f8f9fa; font-weight: 600; color: #555; }}
        tr:hover {{ background: #f8f9fa; }}
        a {{ color: #0066cc; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        .btn {{ background: #0066cc; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; }}
        .btn:hover {{ background: #0052a3; }}
        input[type="file"] {{ margin: 10px 0; padding: 10px; border: 2px dashed #ddd; border-radius: 4px; width: 100%; box-sizing: border-box; }}
        .size {{ color: #666; font-size: 0.9em; }}
        .time {{ color: #999; font-size: 0.85em; }}
    </style>
</head>
<body>
    <h1>📁 OpenClaw Workspace</h1>
    
    <div class="upload-box">
        <h2>📤 上传文件</h2>
        <form enctype="multipart/form-data" method="post">
            <input type="file" name="file" multiple>
            <br><br>
            <input type="submit" value="上传文件" class="btn">
        </form>
    </div>
    
    <div class="file-list">
        <h2>📂 文件列表</h2>
        <table>
            <tr>
                <th>文件名</th>
                <th>大小</th>
                <th>修改时间</th>
            </tr>
"""
            # List files
            for item in sorted(UPLOAD_DIR.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True):
                if item.is_file():
                    stat = item.stat()
                    size = stat.st_size
                    if size > 1024*1024:
                        size_str = f"{size/1024/1024:.2f} MB"
                    elif size > 1024:
                        size_str = f"{size/1024:.2f} KB"
                    else:
                        size_str = f"{size} B"
                    
                    from datetime import datetime
                    mtime = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                    
                    html += f"""
            <tr>
                <td><a href="{item.name}">{item.name}</a></td>
                <td class="size">{size_str}</td>
                <td class="time">{mtime}</td>
            </tr>
"""
            
            html += """
        </table>
    </div>
</body>
</html>
"""
            self.wfile.write(html.encode('utf-8'))
            
        elif path.exists() and path.is_file():
            self.send_response(200)
            # Guess content type
            import mimetypes
            content_type, _ = mimetypes.guess_type(str(path))
            if content_type:
                self.send_header('Content-type', content_type)
            else:
                self.send_header('Content-type', 'application/octet-stream')
            self.send_header('Content-Disposition', f'attachment; filename="{path.name}"')
            self.end_headers()
            
            with open(path, 'rb') as f:
                self.wfile.write(f.read())
        else:
            self.send_error(404, 'File not found')
    
    def do_POST(self):
        """Handle file uploads"""
        content_type = self.headers.get('Content-Type')
        if not content_type or 'multipart/form-data' not in content_type:
            self.send_error(400, 'Invalid content type')
            return
        
        # Parse multipart form data
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={'REQUEST_METHOD': 'POST'}
        )
        
        if 'file' not in form:
            self.send_error(400, 'No file uploaded')
            return
        
        file_item = form['file']
        if file_item.filename:
            # Save file
            filepath = UPLOAD_DIR / file_item.filename
            with open(filepath, 'wb') as f:
                f.write(file_item.file.read())
            
            print(f"Uploaded: {file_item.filename} -> {filepath}")
            
            # Redirect back to file list
            self.send_response(303)
            self.send_header('Location', '/')
            self.end_headers()
        else:
            self.send_error(400, 'No file selected')

if __name__ == '__main__':
    with socketserver.TCPServer(("0.0.0.0", PORT), UploadHandler) as httpd:
        print(f"🚀 File server running at http://0.0.0.0:{PORT}")
        print(f"📁 Upload directory: {UPLOAD_DIR}")
        print("Press Ctrl+C to stop")
        httpd.serve_forever()
