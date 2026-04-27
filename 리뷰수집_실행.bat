@echo off
chcp 65001 > nul
title 코에르 리뷰 수집 프로그램

echo.
echo ┌──────────────────────────────────────────────┐
echo │      코에르 리뷰 수집 프로그램               │
echo │  욕실매트 S - 네이버 브랜드스토어 리뷰 수집  │
echo └──────────────────────────────────────────────┘
echo.

cd /d "%~dp0"

echo  실행 중... 브라우저가 자동으로 열립니다.
echo  완료되면 이 창이 닫히지 않고 결과를 표시합니다.
echo.

node collect_reviews.js

echo.
echo  완료되었습니다. 이 창을 닫으려면 아무 키나 누르세요.
pause > nul
