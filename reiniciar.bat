@echo off
echo ========================================================
echo   Reiniciando Informes Pro...
echo ========================================================
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8501" ^| findstr "LISTENING"') do (
    echo Cerrando proceso en puerto 8501 (PID %%a)...
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul
echo Iniciando servidor limpio...
python -m streamlit run src/main.py
pause
