@echo off
echo Clearing __pycache__
for /F %%F in ('dir /ad /b') do echo %%F
for /F %%F in ('dir /ad /b') do IF EXIST %%F\__pycache__ RMDIR /S /Q %%F\__pycache__
TIMEOUT 5
