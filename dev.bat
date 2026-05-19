@echo off
title Dev — CCTV Control Panel
echo.
echo  ============================================
echo   CCTV Control Panel  —  Modo DEV (Flask)
echo   Acesse: http://127.0.0.1:5000
echo   Ctrl+C para parar
echo  ============================================
echo.

:: Ativa o ambiente virtual se existir
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)

:: Roda o Flask em modo debug com auto-reload
set FLASK_ENV=development
set FLASK_DEBUG=1
python -c "from app import app; app.run(host='127.0.0.1', port=5000, debug=True, use_reloader=True)"

pause
