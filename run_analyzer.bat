@echo off
chcp 65001 > nul
title โปรแกรมวิเคราะห์ Heating Cycle Report (IEC 60840)
echo ==============================================================================
echo  โปรแกรมวิเคราะห์รอบการทดสอบ Heating Cycle Report (IEC 60840)
echo ==============================================================================
echo.
python "%~dp0cycle_analyzer.py" %*
echo.
pause
