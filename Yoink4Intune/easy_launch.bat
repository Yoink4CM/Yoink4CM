@echo off

cd %~dp0

powershell.exe -executionpolicy bypass -file "yoink4intune.ps1"
