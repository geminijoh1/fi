@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3.11 update_auction_and_express.py 2>nul || python update_auction_and_express.py
pause
