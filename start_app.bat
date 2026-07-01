@echo off
chcp 65001 >nul
title Юнит-экономика автозапчастей
echo ========================================
echo 🚗 Запуск приложения...
echo ========================================
echo.

REM Переход в папку скрипта
cd /d "%~dp0"

REM Проверка наличия venv
if not exist "venv\Scripts\activate.bat" (
    echo ❌ Виртуальное окружение не найдено!
    echo 📦 Создаю виртуальное окружение...
    py -3.12 -m venv venv
    call venv\Scripts\activate.bat
    echo 📥 Устанавливаю зависимости...
    pip install streamlit pandas numpy plotly openpyxl scikit-learn chardet Pillow requests python-dateutil pytz polars duckdb joblib
) else (
    call venv\Scripts\activate.bat
)

echo.
echo ✅ Запуск Streamlit...
echo 🌐 Откроется браузер на http://localhost:8501
echo ⏹️ Для остановки закройте это окно или нажмите Ctrl+C
echo.

REM Открываем браузер через 3 секунды
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:8501"

REM Запускаем Streamlit
python -m streamlit run auto_parts_economy.py --server.headless false

pause
