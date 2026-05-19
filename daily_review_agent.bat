@echo off
echo BAT_START_%date%_%time% > "%TEMP%\coeir_bat_trace.log"
chcp 65001 >nul
cd /d "C:\Users\AWESOMATIC\Desktop\코에르\클로드\리뷰수집프로그램"
echo CD_OK_%CD% >> "%TEMP%\coeir_bat_trace.log"
if not exist "logs" mkdir "logs"
set "TS=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%"
set "TS=%TS: =0%"
set "LOG=logs\agent_run_%TS%.log"
type daily_review_prompt.txt | "C:\Users\AWESOMATIC\.local\bin\claude.exe" --print --add-dir "C:\Users\AWESOMATIC\Desktop\코에르\클로드\리뷰수집프로그램" --dangerously-skip-permissions --max-turns 200 > "%LOG%" 2>&1
exit /b %errorlevel%
