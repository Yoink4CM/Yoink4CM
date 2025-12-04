@echo off

cd %~dp0

powershell.exe -executionpolicy bypass -file "yoink4cm_package_cleanup.ps1"

