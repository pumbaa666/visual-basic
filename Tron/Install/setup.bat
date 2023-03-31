@echo off
cls
if %OS% == Windows_NT goto NT

echo.
xcopy "Tron.exe" %windir%"\Menu D‚marrer\Programmes\D‚marrage\Tron.exe"

echo.
xcopy "Tron.exe" %windir%"\Menu D‚marrer\Programmes\Tron.exe"
goto fin

:NT

echo.
xcopy "Tron.exe" "C:\Documents and Settings\All Users\Menu D‚marrer\Programmes\D‚marrage\Tron.exe"

echo.
xcopy "Tron.exe" "C:\Documents and Settings\All Users\Menu D‚marrer\Programmes\Tron.exe"

:fin

echo.
xcopy "vb6fr.dll" %windir%"\system32\vb6fr.dll"/o

echo.
xcopy "Comdlg32.ocx" %windir%"\system\Comdlg32.ocx"/o

echo.
xcopy "RICHTX32.OCX" %windir%"\system\RICHTX32.OCX"/o

echo.
xcopy "MSWINSCK.OCX" %windir%"\system\MSWINSCK.OCX"/o

echo L'installation est termin,e.
pause
