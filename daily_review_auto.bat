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

REM -- Step 1: collect (direct, foreground; 1 retry on transient failure) --
set "TRY=1"
:collect
echo [%date% %time%] Step1 collect (try %TRY%) >> "%LOG%"
node collect_pending.js --scheduled >> "%LOG%" 2>&1

REM -- decide: SKIP / ZERO / GO / STALE / ERR  (STALE = not today = collection failed) --
set "STATUS=GO"
for /f "usebackq delims=" %%i in (`node collect_status.js`) do set "STATUS=%%i"
echo [%date% %time%] status=%STATUS% >> "%LOG%"

REM -- transient failure (STALE/ERR): wait 30s and retry once --
if "%TRY%"=="1" if "%STATUS%"=="STALE" ( set "TRY=2" & echo [%date% %time%] retry collect after 30s >> "%LOG%" & ping -n 31 127.0.0.1 >nul & goto collect )
if "%TRY%"=="1" if "%STATUS%"=="ERR"   ( set "TRY=2" & echo [%date% %time%] retry collect after 30s >> "%LOG%" & ping -n 31 127.0.0.1 >nul & goto collect )

REM -- still failed after retry: skip (avoid reprocessing old data); watchdog will alert --
if "%STATUS%"=="SKIP" goto watchdog
if "%STATUS%"=="ZERO" goto watchdog
if "%STATUS%"=="STALE" goto watchdog
if "%STATUS%"=="ERR" goto watchdog

REM -- Step 0.5: knowledge dump for agent --
echo [%date% %time%] knowledge dump >> "%LOG%"
node dump_knowledge.js > knowledge_dump.txt 2>> "%LOG%"

REM -- remove stale files so a failed agent run cannot post yesterday's data --
if exist replies.json del /f replies.json
if exist judgements.json del /f judgements.json

REM -- Step 2: agent writes judgements.json (minimal fields only, no node exec) --
echo [%date% %time%] Step2 agent judge >> "%LOG%"
type daily_judge_prompt.txt | "%CLAUDE%" --print --add-dir "%DIR%" --dangerously-skip-permissions --max-turns 120 >> "%LOG%" 2>&1

REM -- Step 2.5: merge judgements.json + pending_reviews.json -> full replies.json --
echo [%date% %time%] Step2.5 merge >> "%LOG%"
node merge_replies.js >> "%LOG%" 2>&1

if not exist replies.json (
  echo [%date% %time%] no replies.json - skip posting >> "%LOG%"
  goto watchdog
)

REM -- Step 3: post (--scheduled = auto re-login if session expired) --
echo [%date% %time%] Step3 post >> "%LOG%"
node post_replies.js --scheduled >> "%LOG%" 2>&1

REM -- Step 3.5: verify (--scheduled = auto re-login) --
if exist verify_predictions.js (
  echo [%date% %time%] Step3.5 verify >> "%LOG%"
  node verify_predictions.js --scheduled >> "%LOG%" 2>&1
)

REM -- Step 4: report + slack (direct) --
echo [%date% %time%] Step4 report >> "%LOG%"
node generate_report.js >> "%LOG%" 2>&1

:watchdog
echo [%date% %time%] watchdog >> "%LOG%"
node post_run_check.js >> "%LOG%" 2>&1

echo [%date% %time%] ===== END ===== >> "%LOG%"
