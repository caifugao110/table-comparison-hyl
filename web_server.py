#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import http.server
import socketserver
import os
import json
import subprocess
import datetime
import tempfile
import shutil
from urllib.parse import urlparse, parse_qs

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

# 核心Python脚本路径
CORE_SCRIPT = os.path.join(PROJECT_ROOT, "compare_excel_销售毛利分析表.py")

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
        
        # 设置响应头
        self.send_response(200)
        
        # 设置CORS头
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        
        # 设置正确的Content-Type，确保HTML文件使用UTF-8编码
        if self.path.endswith('.html'):
            self.send_header('Content-Type', 'text/html; charset=utf-8')
        elif self.path.endswith('.css'):
            self.send_header('Content-Type', 'text/css; charset=utf-8')
        elif self.path.endswith('.js'):
            self.send_header('Content-Type', 'application/javascript; charset=utf-8')
        
        self.end_headers()
        
        # 调用父类的do_GET处理静态文件
        super().do_GET()
    
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
            
            # 构建临时目录结构 - 优化：使用更高效的目录创建方式
            temp_dir = tempfile.mkdtemp()
            temp_my_dir = os.path.join(temp_dir, "my")
            temp_from_dir = os.path.join(temp_dir, "from")
            
            # 一次性创建所有目录
            os.makedirs(temp_my_dir, exist_ok=True)
            os.makedirs(temp_from_dir, exist_ok=True)
            
            # 目标文件路径
            temp_baseline = os.path.join(temp_my_dir, "销售毛利分析表.xlsx")
            temp_compare = os.path.join(temp_from_dir, "销售毛利分析表.xlsx")
            
            # 改回使用shutil.copy2，避免原始文件被意外删除
            # 对于小文件，copy2的性能影响可以忽略
            shutil.copy2(baseline_file_path, temp_baseline)
            shutil.copy2(compare_file_path, temp_compare)
            
            # 优化：使用python -u禁用输出缓冲，更快获取脚本输出
            command = [
                'python',
                '-u',  # 禁用输出缓冲
                CORE_SCRIPT
            ]
            
            print(f"执行Python命令: {' '.join(command)}")
            
            # 优化：使用更高效的subprocess调用选项
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                cwd=temp_dir,
                timeout=300,  # 5分钟超时
                bufsize=0,  # 无缓冲I/O
                universal_newlines=True  # 与text=True功能相同，但明确指定
            )
            
            # 优化：直接根据已知的文件名格式生成结果文件路径，避免遍历目录
            original_filename = os.path.basename(baseline_file_path).replace('.xlsx', '')
            result_files = []
            
            # 直接生成预期的结果文件路径
            expected_files = [
                f"{original_filename}_差异结果_{timestamp}.xlsx",
                f"{original_filename}_my_比较结果_{timestamp}.xlsx",
                f"{original_filename}_from_比较结果_{timestamp}.xlsx"
            ]
            
            # 检查文件是否存在
            for expected_file in expected_files:
                file_path = os.path.join(RESULTS_FOLDER, expected_file)
                if os.path.exists(file_path):
                    result_files.append(file_path)
            
            # 如果直接生成的文件路径没有找到，再使用遍历方式兜底
            if not result_files and os.path.exists(RESULTS_FOLDER):
                for f in os.listdir(RESULTS_FOLDER):
                    if f.endswith('.xlsx') and timestamp in f:
                        result_files.append(os.path.join(RESULTS_FOLDER, f))
                        if '差异结果' in f:
                            result_files = [result_files[-1]] + result_files[:-1]
            
            # 构造响应
            response = {
                'success': result.returncode == 0,
                'message': '比较完成' if result.returncode == 0 else '比较失败',
                'resultFiles': result_files,
                'stdout': result.stdout,
                'stderr': result.stderr
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
            # 清理临时文件和目录
            # 注意：使用shutil.copy2时，原始文件仍然存在，需要单独删除
            if 'baseline_file_path' in locals() and os.path.exists(baseline_file_path):
                os.unlink(baseline_file_path)
            if 'compare_file_path' in locals() and os.path.exists(compare_file_path):
                os.unlink(compare_file_path)
            if 'temp_dir' in locals() and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)  # 添加ignore_errors=True，避免删除失败

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