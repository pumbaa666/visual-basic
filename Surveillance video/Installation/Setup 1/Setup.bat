@echo off
if %OS% == Windows_NT goto NT
copy "vb6fr.dll" "c:\Windows\system32\vb6fr.dll"
copy "mscomctl.ocx" "c:\Windows\system32\mscomctl.ocx"
goto end
:nt
copy "vb6fr.dll" "c:\Winnt\System32\vb6fr.dll"
copy "mscomctl.ocx" "c:\Winnt\System32\mscomctl.ocx"
:end

