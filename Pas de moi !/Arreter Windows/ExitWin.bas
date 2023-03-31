Attribute VB_Name = "Module1"
Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" ( _
ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long

