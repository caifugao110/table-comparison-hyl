#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import os
import tempfile
import datetime
import shutil
import webbrowser
import threading

# 导入核心比较函数
from compare_excel_web import compare_excel_files

# 初始化FastAPI应用
app = FastAPI(
    title="Excel比较工具API",
    description="高效的Excel文件比较服务",
    version="1.0.0"
)

# 添加CORS中间件
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 定义结果文件夹
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
# 指向项目根目录的results文件夹
RESULTS_FOLDER = os.path.join(os.path.dirname(PROJECT_ROOT), "results")
os.makedirs(RESULTS_FOLDER, exist_ok=True)

# 挂载静态文件到/static路径
app.mount("/static", StaticFiles(directory=PROJECT_ROOT), name="static")

# 提供主页面路由
@app.get("/")
async def root():
    return FileResponse(os.path.join(PROJECT_ROOT, "index.html"))

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """下载结果文件"""
    try:
        # 构建完整的文件路径
        file_path = os.path.join(RESULTS_FOLDER, filename)
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="文件不存在")
        
        # 返回文件下载响应
        return FileResponse(file_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/compare")
async def compare_excel(
    baselineFile: UploadFile = File(...),
    compareFile: UploadFile = File(...)
):
    """比较两个Excel文件"""
    try:
        # 生成唯一的文件名和时间戳
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        original_filename = os.path.splitext(baselineFile.filename)[0]
        
        # 保存上传的文件到临时位置
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_baseline:
            temp_baseline.write(await baselineFile.read())
            baseline_file_path = temp_baseline.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_compare:
            temp_compare.write(await compareFile.read())
            compare_file_path = temp_compare.name
        
        # 生成结果文件路径
        result_baseline = os.path.join(RESULTS_FOLDER, f"{original_filename}_my_比较结果_{timestamp}.xlsx")
        result_compare = os.path.join(RESULTS_FOLDER, f"{original_filename}_from_比较结果_{timestamp}.xlsx")
        
        # 直接调用核心比较函数，使用多进程执行以提高性能
        import io
        import sys
        from contextlib import redirect_stdout
        
        f = io.StringIO()
        with redirect_stdout(f):
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
        
        # 生成差异结果文件路径
        diff_file = os.path.join(RESULTS_FOLDER, f"{original_filename}_差异结果_{timestamp}.xlsx")
        
        # 收集所有结果文件
        result_files = []
        expected_files = [diff_file, result_baseline, result_compare]
        for expected_file in expected_files:
            if os.path.exists(expected_file):
                result_files.append(expected_file)
        
        # 清理临时文件
        os.unlink(baseline_file_path)
        os.unlink(compare_file_path)
        
        # 返回结果
        return JSONResponse({
            "success": True,
            "message": "比较完成",
            "resultFiles": result_files,
            "stdout": stdout,
            "stderr": ""
        })
        
    except Exception as e:
        # 清理临时文件
        if 'baseline_file_path' in locals() and os.path.exists(baseline_file_path):
            os.unlink(baseline_file_path)
        if 'compare_file_path' in locals() and os.path.exists(compare_file_path):
            os.unlink(compare_file_path)
        
        raise HTTPException(status_code=500, detail=str(e))

def open_browser():
    """延迟打开浏览器，确保服务器已经启动"""
    import time
    time.sleep(2)  # 延迟2秒，等待服务器启动
    webbrowser.open("http://localhost:8000")

if __name__ == "__main__":
    import uvicorn
    # 在单独的线程中打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()
    # 启动服务器
    uvicorn.run(app, host="0.0.0.0", port=8000)
