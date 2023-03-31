VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FrmMain.Visible = False
    Open "c:\temp\rapport.txt" For Append As #1
    Print #1, "Allumage : " & Date & "  :  " & Time
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open "c:\temp\rapport.txt" For Append As #1
    Print #1, "Extinction : " & Date & "  :  " & Time
    Print #1, ""
    Close #1
End Sub
