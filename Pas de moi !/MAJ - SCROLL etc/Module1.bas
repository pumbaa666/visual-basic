Attribute VB_Name = "Module1"
Global Old
'API nécessaire pour le mode "toujours visible"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                ByVal hWndInsertAfter _
                                As Long, ByVal x _
                                As Long, ByVal y _
                                As Long, ByVal cx _
                                As Long, ByVal cy _
                                As Long, ByVal wFlags _
                                As Long) As Long
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
'Etat des touches
Private Declare Function GetKeyState Lib "user32" (ByVal iVirtualKey As Integer) As Long
Global Touches(0 To 255) As Byte
Public Function LireEtatMaj()
    If (&H1 And GetKeyState(vbKeyCapital)) <> 0 Then LireEtatMaj = 1
End Function

Public Function LireEtatNum()
    If (&H1 And GetKeyState(vbKeyNumlock)) <> 0 Then LireEtatNum = 4
End Function
Public Function LireEtatScroll()
    If (&H1 And GetKeyState(vbKeyScrollLock)) <> 0 Then LireEtatScroll = 8
End Function
Public Function LireEtatInsert()
    If (&H1 And GetKeyState(vbKeyInsert)) <> 0 Then LireEtatInsert = 2
End Function
Public Sub ChangerEtat(tc As Long)
Dim RetVal As Long
RetVal = GetKeyboardState(Touches(0))
Touches(tc) = IIf(Touches(tc) = 1, 0, 1)
RetVal = SetKeyboardState(Touches(0))
End Sub

'toujours visible
Public Function Forward(who As Form) 'who correspond au nom de la form  | exemple: form1
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(who.hwnd, -1, 0, 0, 0, 0, Flags)
End Function

'annuler toujours visible
Public Function Backward(who As Form)
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(who.hwnd, -2, 0, 0, 0, 0, Flags)
End Function


