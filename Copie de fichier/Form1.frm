VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim vText As String

    Close #2
    Close #1
    Open "g:\00000001.TMP" For Input As #1
    Open "c:\00000001.TMP" For Output As #2
    Do
        Line Input #1, vText
        Print #2, vText
    Loop While (Not EOF(1))
    Close #2
    Close #1
    MsgBox "ok"
    End
End Sub
