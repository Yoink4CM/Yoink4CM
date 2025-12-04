@echo off

cd %~dp0

powershell.exe -executionpolicy bypass -file "yoink4cm_application_cleanup.ps1"

