VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "a" Then
        KeyAscii = Asc("@")
    ElseIf Chr(KeyAscii) = "b" Then
        KeyAscii = Asc("8")
    ElseIf Chr(KeyAscii) = "c" Then
        KeyAscii = Asc("¢")
    ElseIf Chr(KeyAscii) = "e" Then
        KeyAscii = Asc("€")
    ElseIf Chr(KeyAscii) = "h" Then
        KeyAscii = Asc("#")
    ElseIf Chr(KeyAscii) = "i" Then
        KeyAscii = Asc("1")
    ElseIf Chr(KeyAscii) = "l" Then
        KeyAscii = Asc("£")
    ElseIf Chr(KeyAscii) = "o" Then
        KeyAscii = Asc("0")
    ElseIf Chr(KeyAscii) = "s" Then
        KeyAscii = Asc("5")
    End If
End Sub
