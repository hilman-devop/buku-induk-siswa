@echo off
cd /d "%~dp0"
echo Buku Induk Siswa -> http://127.0.0.1:5055   (admin / admin123)
echo Tutup jendela ini atau tekan Ctrl+C untuk menghentikan server.
".venv\Scripts\python.exe" app.py --port 5055
pause
