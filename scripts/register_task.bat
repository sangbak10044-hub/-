@echo off
REM 이 배치파일을 더블클릭하면, "다우닝대시보드자동갱신"이라는 이름으로
REM Windows 작업 스케줄러에 매시간 실행되는 작업이 등록됩니다.
REM 실행 전에 반드시:
REM   1) scripts\watch_config.example.json 을 scripts\watch_config.json 으로 복사한 뒤
REM      watch_folder / repo_dir 경로를 본인 컴퓨터에 맞게 수정
REM   2) pip install -r requirements.txt (pandas, openpyxl 설치)
REM 를 먼저 해두세요.

setlocal
set SCRIPT_DIR=%~dp0
set CONFIG_PATH=%SCRIPT_DIR%watch_config.json

if not exist "%CONFIG_PATH%" (
  echo [오류] %CONFIG_PATH% 파일이 없습니다.
  echo watch_config.example.json 을 watch_config.json 으로 복사하고 경로를 수정한 뒤 다시 실행하세요.
  pause
  exit /b 1
)

schtasks /create /f /tn "다우닝대시보드자동갱신" ^
  /tr "python \"%SCRIPT_DIR%watch_and_update.py\" --config \"%CONFIG_PATH%\"" ^
  /sc hourly /mo 1 /st 09:00

if %errorlevel% equ 0 (
  echo.
  echo 등록 완료! 매시간 정각 기준으로 %SCRIPT_DIR%watch_and_update.py 가 자동 실행됩니다.
  echo 작업 스케줄러(taskschd.msc)에서 "다우닝대시보드자동갱신" 작업을 확인/수정할 수 있어요.
  echo 지금 바로 한 번 테스트하려면: python "%SCRIPT_DIR%watch_and_update.py" --config "%CONFIG_PATH%"
) else (
  echo [오류] 작업 등록에 실패했습니다. 관리자 권한으로 다시 실행해보세요.
)
pause
