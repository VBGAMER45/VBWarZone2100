@echo off
rem Starts the game in bin\ from wherever this folder is (the same as double-clicking bin\Warzone.exe).
rem Arguments are passed on, e.g. Warzone.cmd --configdir=C:\some\folder (the folder must exist).
start "" /D "%~dp0bin" "%~dp0bin\Warzone.exe" %*
