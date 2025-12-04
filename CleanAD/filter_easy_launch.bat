@echo off

cd %~dp0

powershell.exe -executionpolicy bypass -file "filter_old_pcs.ps1"

