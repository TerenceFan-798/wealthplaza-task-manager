# -*- coding: utf-8 -*-
"""极简静态文件服务（带 CORS），供页面内 fetch 取图。

用法：
  python tools/asset_server.py                 # 默认服务 ../assets，端口 8765
  python tools/asset_server.py <目录> <端口>
"""
import http.server
import os
import socketserver
import sys

DEFAULT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
DIR = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DIR
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8765


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, *a):
        pass


if not os.path.isdir(DIR):
    print(f"目录不存在：{DIR}", flush=True)
    sys.exit(1)

os.chdir(DIR)
socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
    print(f"serving {DIR} on 127.0.0.1:{PORT}", flush=True)
    httpd.serve_forever()
