Attribute VB_Name = "Module1"

Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2

Sub w95shutdown()
    'arr�t du systeme
    R& = ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub

Sub w95reboot()
    'arr�t et red�marrage du systeme
    R& = ExitWindowsEx(EWX_REBOOT, 0)
End Sub
