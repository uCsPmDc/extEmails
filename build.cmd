@echo off
rem Call pyinstaller to create an EXE from a Python script
rem with a custom icon. If created, make a release folder
rem and copy the documentation and exe to it.

pyinstaller --onefile --icon=./assets/icon.ico extEmails.py

if not exist ".\release" md ".\release" >nul

if exist ".\dist\extEmails.exe" (
    if exist ".\readme.pdf" copy ".\readme.pdf" ".\release\extEmails.pdf" > nul
    copy ".\dist\extEmails.exe" ".\release\extEmails.exe" > nul
    copy ".\dist\extEmails.exe" ".\release\extEmails.exe" > nul
)