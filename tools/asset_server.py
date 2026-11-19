# -*- coding: utf-8 -*-
"""极简静态文件服务，带 CORS，供页面内 fetch 取图使用。"""
import http.server, socketserver, functools, os

DIR = r"C:\Users\Vincent\WorkBuddy\2026-11-18-16-08-38\assets"
PORT = 8765


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, *a):
        pass


os.chdir(DIR)
with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
    print("serving on", PORT, flush=True)
    httpd.serve_forever()
