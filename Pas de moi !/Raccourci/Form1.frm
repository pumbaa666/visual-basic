VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cliquez ici"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CmdD 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim Dossier_Bureau As String
    On Error Resume Next
    Dossier_Bureau = GetDesktopPath$
    CmdD.ShowOpen
    Call OSfCreateShellLink(Dossier_Bureau, "LINK", CmdD.FileName, "", 0, "")
End Sub


