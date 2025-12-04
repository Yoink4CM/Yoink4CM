@echo off

cd %~dp0

powershell.exe -executionpolicy bypass -file "document_old_pcs.ps1"
