"""启动脚本 - 开发模式使用 Flask 内置服务器"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app

if __name__ == '__main__':
    print("=" * 50)
    print("  补单汇总工具 - Web 版")
    print(f"  地址: http://localhost:5000")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=True)
