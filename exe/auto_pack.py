#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel文件比较工具自动化打包脚本
功能：
1. 将compare_excel.py和图标文件打包成不带控制台的EXE文件
2. 生成文件：EXCEL文件比较工具.exe
3. 清理打包过程中的临时文件
"""

import os
import shutil
import subprocess
import sys

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 输入文件路径
SCRIPT_PATH = os.path.join(PROJECT_ROOT, "gui", "compare_excel.py")
ICON_PATH = os.path.join(PROJECT_ROOT, "ico", "compare_excel.ico")

# 输出目录
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "exe")

# 要删除的临时文件/目录
TEMP_FILES = [
    os.path.join(PROJECT_ROOT, "build"),
    os.path.join(PROJECT_ROOT, "EXCEL文件比较工具.spec"),
    os.path.join(PROJECT_ROOT, "ExcelCompare.spec")  # 旧版spec文件
]


def run_command(cmd, cwd=None):
    """执行命令并返回结果"""
    print(f"执行命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    print(f"返回码: {result.returncode}")
    if result.stdout:
        print(f"标准输出: {result.stdout}")
    if result.stderr:
        print(f"标准错误: {result.stderr}")
    return result


def clean_temp_files():
    """清理临时文件和目录"""
    print("\n=== 清理临时文件 ===")
    for item in TEMP_FILES:
        if os.path.exists(item):
            if os.path.isdir(item):
                print(f"删除目录: {item}")
                shutil.rmtree(item, ignore_errors=True)
            else:
                print(f"删除文件: {item}")
                os.remove(item)
        else:
            print(f"跳过不存在的项: {item}")


def main():
    """主函数"""
    print("=== Excel文件比较工具自动化打包脚本 ===")
    
    # 1. 检查输入文件是否存在
    print("\n=== 检查输入文件 ===")
    if not os.path.exists(SCRIPT_PATH):
        print(f"错误: 脚本文件不存在 - {SCRIPT_PATH}")
        sys.exit(1)
    
    if not os.path.exists(ICON_PATH):
        print(f"错误: 图标文件不存在 - {ICON_PATH}")
        sys.exit(1)
    
    # 2. 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"输出目录: {OUTPUT_DIR}")
    
    # 3. 执行PyInstaller打包
    print("\n=== 执行打包命令 ===")
    # 直接使用pyinstaller命令，而不是通过python -m pyinstaller
    pyinstaller_cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        f"--icon={ICON_PATH}",
        "--name", "EXCEL文件比较工具",
        f"--distpath={OUTPUT_DIR}",
        f"--add-data={os.path.join(PROJECT_ROOT, 'ico')};ico",
        SCRIPT_PATH
    ]
    
    result = run_command(pyinstaller_cmd, cwd=PROJECT_ROOT)
    if result.returncode != 0:
        print("打包失败!")
        sys.exit(1)
    
    # 4. 清理临时文件
    clean_temp_files()
    
    # 5. 验证输出文件
    print("\n=== 验证输出文件 ===")
    output_exe = os.path.join(OUTPUT_DIR, "EXCEL文件比较工具.exe")
    if os.path.exists(output_exe):
        print(f"✅ 打包成功!")
        print(f"输出文件: {output_exe}")
        print(f"文件大小: {os.path.getsize(output_exe) / 1024 / 1024:.2f} MB")
    else:
        print(f"❌ 打包失败，输出文件不存在: {output_exe}")
        sys.exit(1)
    
    print("\n🎉 自动化打包流程完成!")


if __name__ == "__main__":
    main()
