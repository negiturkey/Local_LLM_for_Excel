import http.server
import ssl
import os
import subprocess
import urllib.request
import json

# 設定
PORT = 3000
CERT_FILE = "server/cert.pem"
KEY_FILE = "server/key.pem"

# バックエンドの設定
BACKENDS = {
    "ollama": "http://127.0.0.1:11434",
    "lmstudio": "http://127.0.0.1:1234" 
}

def generate_cert():
    if not os.path.exists(CERT_FILE):
        print("Creating self-signed certificate...")
        try:
            config_paths = [
                r"C:\Program Files\Git\usr\ssl\openssl.cnf",
                r"C:\Program Files\OpenSSL-Win64\bin\openssl.cnf"
            ]
            config_path = next((p for p in config_paths if os.path.exists(p)), None)
            
            cmd = [
                "openssl", "req", "-x509", "-newkey", "rsa:2048", 
                "-keyout", KEY_FILE, "-out", CERT_FILE, 
                "-days", "365", "-nodes", 
                "-subj", "/CN=localhost"
            ]
            if config_path: cmd.extend(["-config", config_path])
            
            subprocess.run(cmd, check=True)
            print("Certificate created.")
        except Exception as e:
            print(f"OpenSSL failed: {e}")

class MyHandler(http.server.SimpleHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        if self.path == '/api/env':
            self.handle_env_request()
        elif self.path.startswith('/api/proxy/'):
            self.proxy_request('GET')
        elif self.path == '/api/templates':
            self.handle_templates_get()
        else:
            super().do_GET()

    def handle_env_request(self):
        api_key = None
        env_path = ".env"
        if os.path.exists(env_path):
            try:
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("GEMINI_API_KEY="):
                            api_key = line.split("=", 1)[1].strip().strip('"').strip("'")
                            break
            except Exception as e:
                print(f"Error reading .env: {e}")
        
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"apiKey": api_key}).encode("utf-8"))

    def do_POST(self):
        if self.path.startswith('/api/proxy/'):
            self.proxy_request('POST')
        elif self.path == '/api/templates':
            self.handle_templates_post()
        else:
            self.send_error(404, "Not Found")

    def handle_templates_get(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        template_file = "src/user_templates.json"
        if os.path.exists(template_file):
            try:
                with open(template_file, "r", encoding="utf-8") as f:
                    self.wfile.write(f.read().encode("utf-8"))
            except Exception as e:
                self.wfile.write(b"{}")
        else:
            self.wfile.write(b"{}")

    def handle_templates_post(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        template_file = "src/user_templates.json"
        try:
            # Validate JSON
            json_data = json.loads(body)
            with open(template_file, "w", encoding="utf-8") as f:
                json.dump(json_data, f, indent=4, ensure_ascii=False)
            
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{"status": "ok"}')
        except Exception as e:
            self.send_error(500, str(e))

    def proxy_request(self, method):
        parts = self.path.split('/')
        if len(parts) < 5:
            self.send_error(400, "Invalid path")
            return
            
        backend_name = parts[3]
        if backend_name not in BACKENDS:
            self.send_error(404, "Backend not found")
            return
            
        target_url = f"{BACKENDS[backend_name]}/{'/'.join(parts[4:])}"
        print(f"[Proxy] {method} -> {target_url}")
        
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length) if content_length > 0 else None
        
        try:
            req = urllib.request.Request(target_url, data=body, method=method)
            if self.headers.get('Content-Type'):
                req.add_header('Content-Type', self.headers.get('Content-Type'))
            
            # タイムアウトを1200秒（20分）に延長（70GB超のモデルをHDDからロードする場合を考慮）
            with urllib.request.urlopen(req, timeout=1200) as response:
                self.send_response(response.status)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(response.read())
        except Exception as e:
            print(f"[Error] {e}")
            self.send_error(500, str(e))

def run_server():
    base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(base_path)
    generate_cert()
    # 全てのインターフェースで待ち受け（localhost / 127.0.0.1 両対応）
    httpd = http.server.HTTPServer(('', PORT), MyHandler)
    # SSL無効化 (証明書エラー回避用)
    # if os.path.exists(CERT_FILE):
    #     context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    #     context.load_cert_chain(certfile=CERT_FILE, keyfile=KEY_FILE)
    #     httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
    #     print(f"HTTPS Server: https://127.0.0.1:{PORT}")
    # else:
    print(f"HTTP Server: http://127.0.0.1:{PORT}")

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        httpd.server_close()

if __name__ == "__main__":
    run_server()
