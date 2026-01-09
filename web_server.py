#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import http.server
import socketserver
import os
import json
import datetime
import tempfile
import shutil
from urllib.parse import urlparse, parse_qs

# 直接导入核心比较函数，避免外部进程调用
from compare_excel_web import compare_excel_files

# 替代cgi模块的简单multipart解析器

def parse_multipart(fp, boundary):
    """优化的multipart/form-data解析器，使用更高效的方式处理文件"""
    # 移除seek操作，因为HTTP请求流是不可寻址的
    data = fp.read()
    boundary = boundary.encode()
    parts = data.split(b'--' + boundary)
    
    files = {}
    
    for part in parts[1:-1]:  # 跳过第一个和最后一个边界
        part = part.strip()
        if not part:
            continue
            
        # 分离头部和内容
        header_end = part.find(b'\r\n\r\n')
        if header_end == -1:
            continue
            
        headers = part[:header_end].decode()
        content = part[header_end + 4:]
        
        # 解析文件名和字段名
        name = None
        filename = None
        content_type = None
        
        for line in headers.split('\r\n'):
            if line.startswith('Content-Disposition:'):
                # 提取name和filename
                disp = line.split(':', 1)[1].strip()
                for param in disp.split(';'):
                    param = param.strip()
                    if param.startswith('name='):
                        name = param.split('=', 1)[1].strip('"')
                    elif param.startswith('filename='):
                        filename = param.split('=', 1)[1].strip('"')
            elif line.startswith('Content-Type:'):
                content_type = line.split(':', 1)[1].strip()
        
        if name and filename:
            # 保存文件到临时位置
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as f:
                f.write(content)
                files[name] = {
                    'filename': filename,
                    'path': f.name,
                    'content_type': content_type
                }
    
    return files

# 定义服务器端口
PORT = 8000

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

# 创建results文件夹（如果不存在）
RESULTS_FOLDER = os.path.join(PROJECT_ROOT, "results")
os.makedirs(RESULTS_FOLDER, exist_ok=True)

class RequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        # 处理GET请求
        parsed_path = urlparse(self.path)
        
        # 如果路径是根目录，返回index.html
        if parsed_path.path == '/':
            self.path = '/index.html'
        
        # 调用父类的do_GET处理静态文件
        super().do_GET()
    
    def end_headers(self):
        # 添加CORS头
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        
        # 调用父类的end_headers
        super().end_headers()
    
    def do_POST(self):
        # 处理POST请求
        if self.path == '/api/compare':
            self.handle_compare_request()
        else:
            self.send_error(404, "Not Found")
    
    def do_OPTIONS(self):
        # 处理OPTIONS请求（用于CORS预检）
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def handle_compare_request(self):
        # 解析multipart/form-data请求
        content_type = self.headers['Content-Type']
        if not content_type.startswith('multipart/form-data'):
            self.send_error(400, "Bad Request: Only multipart/form-data is supported")
            return
        
        # 提取boundary
        boundary = content_type.split('boundary=')[1]
        
        # 使用自定义解析器解析multipart数据
        files = parse_multipart(self.rfile, boundary)
        
        # 检查是否有文件字段
        if 'baselineFile' not in files or 'compareFile' not in files:
            self.send_error(400, "Bad Request: Missing file fields")
            return
        
        # 获取文件路径
        baseline_file_path = files['baselineFile']['path']
        compare_file_path = files['compareFile']['path']
        
        try:
            # 生成结果文件名和时间戳
            original_filename = os.path.basename(baseline_file_path).replace('.xlsx', '')
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 不需要创建临时目录和复制文件，直接使用上传的文件路径
            
            # 生成结果文件路径
            result_baseline = os.path.join(RESULTS_FOLDER, f"{original_filename}_my_比较结果_{timestamp}.xlsx")
            result_compare = os.path.join(RESULTS_FOLDER, f"{original_filename}_from_比较结果_{timestamp}.xlsx")
            
            print(f"直接调用核心比较函数...")
            
            # 直接调用核心比较函数，避免外部进程开销
            # 重定向stdout以捕获输出
            import io
            import sys
            from contextlib import redirect_stdout
            
            f = io.StringIO()
            with redirect_stdout(f):
                # 直接使用上传的文件路径，无需复制到临时目录
                compare_excel_files(
                    baseline_file_path,  # 基准文件路径
                    compare_file_path,   # 比较文件路径
                    result_baseline,     # 输出基准文件路径
                    result_compare,      # 输出比较文件路径
                    original_filename,   # 原始文件名
                    timestamp            # 时间戳
                )
            
            # 获取函数输出
            stdout = f.getvalue()
            stderr = ""  # 直接调用不会产生stderr
            returncode = 0  # 假设成功，实际应用中需要添加错误处理
            
            # 优化：直接使用已知的结果文件路径，避免遍历目录
            result_files = []
            
            # 直接生成预期的结果文件路径
            diff_file = os.path.join(RESULTS_FOLDER, f"{original_filename}_差异结果_{timestamp}.xlsx")
            expected_files = [
                diff_file,
                result_baseline,
                result_compare
            ]
            
            # 检查文件是否存在
            for expected_file in expected_files:
                if os.path.exists(expected_file):
                    result_files.append(expected_file)
            
            # 构造响应
            response = {
                'success': returncode == 0,
                'message': '比较完成' if returncode == 0 else '比较失败',
                'resultFiles': result_files,
                'stdout': stdout,
                'stderr': stderr
            }
            
            # 发送响应
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response).encode('utf-8'))
            
        except Exception as e:
            print(f"处理请求时出错: {e}")
            
            # 发送错误响应
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            error_response = {
                'success': False,
                'error': str(e)
            }
            self.wfile.write(json.dumps(error_response).encode('utf-8'))
        finally:
            # 清理临时文件
            if 'baseline_file_path' in locals() and os.path.exists(baseline_file_path):
                os.unlink(baseline_file_path)
            if 'compare_file_path' in locals() and os.path.exists(compare_file_path):
                os.unlink(compare_file_path)

# 创建服务器
with socketserver.TCPServer(("", PORT), RequestHandler) as httpd:
    print(f"\n服务器运行在 http://localhost:{PORT}")
    print(f"请在浏览器中访问: http://localhost:{PORT}\n")
    
    try:
        # 启动服务器
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
        httpd.shutdown()