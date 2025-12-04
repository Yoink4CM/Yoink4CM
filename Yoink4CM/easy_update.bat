@echo off

cd %~dp0

powershell.exe -executionpolicy bypass -file "update.ps1"


pause