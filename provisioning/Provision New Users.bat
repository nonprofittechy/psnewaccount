@echo off

Echo Loading the user creation script, please wait...
Echo A File Open dialog box will pop up soon to allow you choose the input file.

%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe  -executionpolicy bypass .\provision.ps1

pause