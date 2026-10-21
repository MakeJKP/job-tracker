@echo off
:: run.bat — double-click this to launch Job Tracker
::            without opening a black console window.
pythonw job_tracker.py
if errorlevel 1 python job_tracker.py
