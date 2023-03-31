@echo off
cls
if %OS% == Windows_NT goto NT
copy "Anniversaire.exe" %windir%"\Menu D‚marrer\Programmes\D‚marrage\anniversaire.exe"
copy "Anniversaire.exe" %windir%"\Menu D‚marrer\Programmes\anniversaire.exe"
goto fin

:NT
xcopy "Anniversaire.exe" "C:\Documents and Settings\All Users\Menu D‚marrer\Programmes\D‚marrage\anniversaire.exe"
xcopy "Anniversaire.exe" "C:\Documents and Settings\All Users\Menu D‚marrer\Programmes\anniversaire.exe"

:fin
xcopy "vb6fr.dll" %windir%"\system32\vb6fr.dll"
xcopy "donnees.dat" "c:\temp\donnees.dat"
echo L'installation est termin,e.
pause