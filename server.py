import http.server
import json
import os
import sys
import webbrowser

PORT = 8080
DATA_FILE = 'data.json'

DEFAULT_DATA = {
    "students": [],
    "lessons": [],
    "attendance": {}
}

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return DEFAULT_DATA.copy()

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/api/data':
            data = load_data()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))
        elif self.path == '/' or self.path == '':
            self.path = '/calendar.html'
            super().do_GET()
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/data':
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length).decode('utf-8')
            data = json.loads(body)
            save_data(data)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        pass

if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    if not os.path.exists(DATA_FILE):
        save_data(DEFAULT_DATA)

    server = http.server.HTTPServer(('', PORT), Handler)
    print(f'Сервер запущен: http://localhost:{PORT}')
    print('Нажмите Ctrl+C для остановки')
    webbrowser.open(f'http://localhost:{PORT}')

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nСервер остановлен')
        server.server_close()
