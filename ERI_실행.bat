@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title ERI 텍스트 복잡도 계산기

rem ---------------------------------------------------------------
rem  ERI 계산기 실행 파일 - 더블클릭하면 실행됩니다.
rem  1) 진짜 파이썬 3.10 이상을 찾습니다. Microsoft Store 바로가기는 건너뜁니다.
rem  2) 처음 한 번은 필요한 라이브러리를 자동으로 설치합니다.
rem  3) run_eri.py 를 실행합니다.
rem ---------------------------------------------------------------

set "PY="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
if defined PY goto found
python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 set "PY=python"
if defined PY goto found
goto nopython

:found
%PY% -c "import kiwipiepy, openpyxl, numpy" >nul 2>&1
if not errorlevel 1 goto run
echo 처음 실행이라 필요한 라이브러리를 설치합니다. 1~2분 정도 걸립니다...
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto pipfail

:run
%PY% run_eri.py %*
if errorlevel 1 pause
exit /b

:nopython
echo.
echo [오류] 파이썬 3.10 이상을 찾지 못했습니다.
echo.
echo  - 'python'을 입력했을 때 'Python'만 나오고 끝난다면, 그것은 진짜 파이썬이 아니라
echo    Windows의 Microsoft Store 바로가기입니다.
echo  - https://www.python.org/downloads/ 에서 파이썬을 설치하세요.
echo    설치 첫 화면에서 "Add python.exe to PATH" 를 꼭 체크하세요.
echo  - 설치 후 이 창을 닫고 ERI_실행.bat 을 다시 더블클릭하세요.
echo.
start "" "https://www.python.org/downloads/"
pause
exit /b 1

:pipfail
echo.
echo [오류] 라이브러리 설치에 실패했습니다. 인터넷 연결을 확인한 뒤 다시 실행하세요.
echo        직접 설치하려면: %PY% -m pip install -r requirements.txt
pause
exit /b 1
