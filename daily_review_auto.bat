@echo off
REM ===== daily review orchestrator (ASCII only, no chcp) =====
REM node steps run directly here (always foreground). Agent only writes replies.json.
REM No "chcp" here: chcp 65001 inside a .bat corrupts cmd file parsing under the scheduler.
REM No literal Korean path: %~dp0 resolves the real (Korean) folder at runtime.

cd /d "%~dp0"
set "DIR=%~dp0"
if "%DIR:~-1%"=="\" set "DIR=%DIR:~0,-1%"
set "CLAUDE=C:\Users\AWESOMATIC\.local\bin\claude.exe"
set "LOG=%DIR%\logs\auto_%date:~0,4%%date:~5,2%%date:~8,2%.log"

echo. >> "%LOG%"
echo [%date% %time%] ===== START ===== >> "%LOG%"

REM -- Step 1: collect (direct, foreground) --
echo [%date% %time%] Step1 collect >> "%LOG%"
node collect_pending.js --scheduled >> "%LOG%" 2>&1

REM -- decide: SKIP / ZERO / GO --
set "STATUS=GO"
for /f "usebackq delims=" %%i in (`node -e "try{var d=require('./pending_reviews.json');process.stdout.write(d.skipped===true?'SKIP':(d.totalReviews>0?'GO':'ZERO'))}catch(e){process.stdout.write('ERR')}"`) do set "STATUS=%%i"
echo [%date% %time%] status=%STATUS% >> "%LOG%"

if "%STATUS%"=="SKIP" goto watchdog
if "%STATUS%"=="ZERO" goto watchdog
if "%STATUS%"=="ERR" goto watchdog

REM -- Step 0.5: knowledge dump for agent --
echo [%date% %time%] knowledge dump >> "%LOG%"
node dump_knowledge.js > knowledge_dump.txt 2>> "%LOG%"

REM -- remove stale replies.json so a failed agent run cannot post yesterday's file --
if exist replies.json del /f replies.json

REM -- Step 2: agent writes replies.json (no node exec) --
echo [%date% %time%] Step2 agent judge >> "%LOG%"
type daily_judge_prompt.txt | "%CLAUDE%" --print --add-dir "%DIR%" --dangerously-skip-permissions --max-turns 120 >> "%LOG%" 2>&1

if not exist replies.json (
  echo [%date% %time%] no replies.json - skip posting >> "%LOG%"
  goto watchdog
)

REM -- Step 3: post (direct) --
echo [%date% %time%] Step3 post >> "%LOG%"
node post_replies.js >> "%LOG%" 2>&1

REM -- Step 3.5: verify (if present) --
if exist verify_predictions.js (
  echo [%date% %time%] Step3.5 verify >> "%LOG%"
  node verify_predictions.js >> "%LOG%" 2>&1
)

REM -- Step 4: report + slack (direct) --
echo [%date% %time%] Step4 report >> "%LOG%"
node generate_report.js >> "%LOG%" 2>&1

:watchdog
echo [%date% %time%] watchdog >> "%LOG%"
node post_run_check.js >> "%LOG%" 2>&1

echo [%date% %time%] ===== END ===== >> "%LOG%"
