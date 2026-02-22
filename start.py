#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepExtract 一键启动脚本
同时启动前端和后端服务
"""

import os
import sys
import subprocess
import webbrowser
from pathlib import Path


def check_python():
    """检查 Python 版本"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("[错误] 需要 Python 3.8 或更高版本")
        return False
    print(f"[✓] Python 版本: {version.major}.{version.minor}.{version.micro}")
    return True


def install_dependencies():
    """安装依赖"""
    backend_dir = Path(__file__).parent / "backend"
    req_file = backend_dir / "requirements.txt"

    if req_file.exists():
        print("[...] 正在安装依赖...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r", str(req_file)],
                check=True,
                capture_output=True,
            )
            print("[✓] 依赖安装完成")
        except subprocess.CalledProcessError:
            print("[!] 依赖安装可能出现问题，继续启动...")
    else:
        print("[!] 未找到 requirements.txt")


def start_service():
    """启动 Flask 服务"""
    backend_dir = Path(__file__).parent / "backend"
    app_file = backend_dir / "app.py"

    if not app_file.exists():
        print(f"[错误] 未找到 {app_file}")
        return False

    print("\n" + "=" * 50)
    print("    DeepExtract - 文档结构化萃取服务")
    print("=" * 50)
    print("\n    访问地址:")
    print("    http://localhost:5000")
    print("\n" + "=" * 50 + "\n")

    # 自动打开浏览器
    webbrowser.open("http://localhost:5000")

    # 启动 Flask
    os.chdir(backend_dir)
    subprocess.run([sys.executable, str(app_file)])

    return True


def main():
    print("\n" + "=" * 50)
    print("    DeepExtract 启动脚本")
    print("=" * 50 + "\n")

    if not check_python():
        input("\n按回车键退出...")
        return

    install_dependencies()
    start_service()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n服务已停止")
        sys.exit(0)
