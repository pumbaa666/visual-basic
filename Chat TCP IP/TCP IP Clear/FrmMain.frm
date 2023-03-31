VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   1620
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock WsClear 
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim vCount As Integer
    For vCount = 0 To 5000
        WsClear.LocalPort = 1001 + vCount
        If WsClear.State <> sckClosed Then
            WsClear.Close
        End If
    Next
    End
End Sub
