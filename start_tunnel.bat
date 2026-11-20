@echo off
title Tunnel
cd /d "%~dp0"
echo Tunnel for http://127.0.0.1:8501
echo Make sure streamlit is running first!
.\cloudflared.exe tunnel --url http://127.0.0.1:8501
pause