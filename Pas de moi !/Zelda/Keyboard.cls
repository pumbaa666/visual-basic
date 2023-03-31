VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "keyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'class de BumpMAN tres utils
Public LeftKey As Boolean
Public RightKey As Boolean
Public UpKey As Boolean
Public DownKey As Boolean
Public F1Key As Boolean
Public F2Key As Boolean
Public ShiftKey As Boolean
Public F3key As Boolean
Public SpaceKey As Boolean
Public ControlKey As Boolean
Public Sub SetKeysDown(KeyCode As Integer)
If KeyCode = vbKeyRight Then RightKey = True
If KeyCode = vbKeyLeft Then LeftKey = True
If KeyCode = vbKeyUp Then UpKey = True
If KeyCode = vbKeyDown Then DownKey = True
If KeyCode = vbKeyF1 Then F1Key = True
If KeyCode = vbKeyF2 Then F2Key = True
If KeyCode = vbKeyShift Then ShiftKey = True
If KeyCode = vbKeyF3 Then F3key = True
If KeyCode = vbKeySpace Then SpaceKey = True
If KeyCode = vbKeyControl Then ControlKey = True
End Sub

Public Sub SetKeysUp(KeyCode As Integer)
If KeyCode = vbKeyRight Then RightKey = False
If KeyCode = vbKeyLeft Then LeftKey = False
If KeyCode = vbKeyUp Then UpKey = False
If KeyCode = vbKeyDown Then DownKey = False
If KeyCode = vbKeyF1 Then F1Key = False
If KeyCode = vbKeyF2 Then F2Key = False
If KeyCode = vbKeyShift Then ShiftKey = False
If KeyCode = vbKeyF3 Then F3key = False
If KeyCode = vbKeySpace Then SpaceKey = False
If KeyCode = vbKeyControl Then ControlKey = False

End Sub
