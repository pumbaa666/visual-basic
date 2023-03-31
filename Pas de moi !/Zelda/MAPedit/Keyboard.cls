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
'class de BumpMAN tres utile
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
Public TKey As Boolean

Public num1Key As Boolean
Public num2Key As Boolean
Public num3Key As Boolean
Public num4Key As Boolean
Public num5Key As Boolean
Public num6Key As Boolean
Public num7Key As Boolean
Public num8Key As Boolean
Public num9Key As Boolean
Public num0Key As Boolean

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
If KeyCode = vbKey1 Then num1Key = True
If KeyCode = vbKey2 Then num2Key = True
If KeyCode = vbKey3 Then num3Key = True
If KeyCode = vbKey4 Then num4Key = True
If KeyCode = vbKey5 Then num5Key = True
If KeyCode = vbKey6 Then num6Key = True
If KeyCode = vbKey7 Then num7Key = True
If KeyCode = vbKey8 Then num8Key = True
If KeyCode = vbKey9 Then num9Key = True
If KeyCode = vbKey0 Then num0Key = True
If KeyCode = vbKeyT Then TKey = True

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
If KeyCode = vbKey1 Then num1Key = False
If KeyCode = vbKey2 Then num2Key = False
If KeyCode = vbKey3 Then num3Key = False
If KeyCode = vbKey4 Then num4Key = False
If KeyCode = vbKey5 Then num5Key = False
If KeyCode = vbKey6 Then num6Key = False
If KeyCode = vbKey7 Then num7Key = False
If KeyCode = vbKey8 Then num8Key = False
If KeyCode = vbKey9 Then num9Key = False
If KeyCode = vbKey0 Then num0Key = False
If KeyCode = vbKeyT Then TKey = False

End Sub
