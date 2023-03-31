VERSION 5.00
Begin VB.Form FrmFichier 
   Caption         =   "Choissez un fichier"
   ClientHeight    =   4320
   ClientLeft      =   360
   ClientTop       =   390
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtFichier 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2400
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "FrmFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    If InStr(1, TxtFichier.Text, "-") < 1 Or InStr(1, TxtFichier.Text, ".txt") < 1 Then
        MsgBox "Fichier invalid", vbCritical, "Erreur"
    Else
        FrmMain.Show
        FrmFichier.Hide
        Chargement
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    TxtFichier.Text = File1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    TxtFichier.Text = Dir1.Path
End Sub

Private Sub File1_Click()
    TxtFichier.Text = File1.Path & "\" & File1.FileName
End Sub

Private Sub File1_DblClick()
    CmdOk_Click
End Sub
