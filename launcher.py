import subprocess
import sys
import os

# 切到当前目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 启动 streamlit
subprocess.run([
    sys.executable, "-m", "streamlit", "run", "app.py",
    "--server.port", "8501",
    "--browser.gatherUsageStats", "false"
])