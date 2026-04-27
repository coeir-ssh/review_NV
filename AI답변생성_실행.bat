@echo off
chcp 65001 > nul
title 코에르 리뷰 AI 답변 생성

echo.
echo ┌─────────────────────────────────────────────────────┐
echo │   코에르 리뷰 AI 답변 생성 프로그램                │
echo │   - 환불 여부 AI 자동 판단                         │
echo │   - 리뷰별 100자 답변 자동 생성                    │
echo └─────────────────────────────────────────────────────┘
echo.

cd /d "%~dp0"

:: config.js에서 API 키 확인
node -e "const c=require('./config'); if(!c.ANTHROPIC_API_KEY){process.exit(1);}" 2>nul
if %errorlevel% neq 0 (
    echo  [설정 필요] Anthropic API 키가 없습니다.
    echo.
    echo  API 키 발급: https://console.anthropic.com
    echo.
    set /p APIKEY="  API 키를 입력하세요 (sk-ant-...): "

    :: config.js 업데이트
    node -e "const fs=require('fs'); const content=fs.readFileSync('config.js','utf8').replace(\"process.env.ANTHROPIC_API_KEY || ''\",\"'%APIKEY%'\"); fs.writeFileSync('config.js',content);"
    echo.
    echo  API 키가 config.js에 저장되었습니다.
    echo.
)

echo  AI 처리 중... (약 3~5분 소요, 잠시 기다려 주세요)
echo.

node generate_responses.js

echo.
echo  완료! 엑셀 파일을 열어 결과를 확인하세요.
echo  이 창을 닫으려면 아무 키나 누르세요.
pause > nul
