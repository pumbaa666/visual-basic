VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Renommeur"
   ClientHeight    =   6855
   ClientLeft      =   465
   ClientTop       =   555
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtChemin 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "Chemin"
      Top             =   360
      Width           =   6015
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "A &propos"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   2895
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Action à effectuer"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "&Rafraichir"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   5640
      Width           =   2895
   End
   Begin VB.ComboBox ComboType 
      Height          =   315
      ItemData        =   "FrmMain.frx":0000
      Left            =   3240
      List            =   "FrmMain.frx":0019
      TabIndex        =   5
      Text            =   "Type de fichier"
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   6240
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   3240
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Séléctionnez le dossier"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAction_Click()
    FrmSyntaxe.Top = 100
    FrmSyntaxe.Left = FrmMain.Left + FrmMain.Width + 100
    FrmSyntaxe.Show
End Sub

Private Sub CmdAbout_Click()
    FrmAbout.Left = FrmMain.Left + FrmMain.Width + 100
    FrmAbout.Top = 100
    FrmAbout.Show
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdRefresh_Click()
    File1.Refresh
End Sub

Private Sub ComboType_Click()
    File1.Pattern = ComboType.Text
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Scroll()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo NoDrive
    Dir1.Path = Drive1.Drive
    Exit Sub
NoDrive:
    MsgBox "Le périphérique " & Drive1.Drive & " n'est pas disponible.", vbCritical, "Erreur"
    Drive1.Drive = "c:"
End Sub

Private Sub File1_PathChange()
    TxtChemin.Text = Dir1.Path
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:"
    FrmMain.Top = 100
    FrmMain.Left = 100

    FrmSyntaxe.Top = 100
    FrmSyntaxe.Left = FrmMain.Left + FrmMain.Width + 100
    FrmSyntaxe.Show

'    FrmAbout.Left = FrmMain.Left + FrmMain.Width + 100
'    FrmAbout.Top = FrmSyntaxe.Top + FrmSyntaxe.Height + 100
'    FrmAbout.Show
End Sub

Private Sub Form_Resize()
    If FrmMain.Height > 4000 Then
        CmdAction.Top = FrmMain.Height - 1950
        CmdRefresh.Top = FrmMain.Height - 1950
        CmdAbout.Top = FrmMain.Height - 1380
        CmdQuitter.Top = FrmMain.Height - 1380
    
        Dir1.Height = FrmMain.Height - 3000
        File1.Height = FrmMain.Height - 3000
    End If
    ComboType.Left = FrmMain.Width / 2
    CmdRefresh.Left = FrmMain.Width / 2
    CmdQuitter.Left = FrmMain.Width / 2
    File1.Left = FrmMain.Width / 2
    
    Dir1.Width = FrmMain.Width / 2 - 300
    File1.Width = FrmMain.Width / 2 - 300
    ComboType.Width = FrmMain.Width / 2 - 300
    CmdRefresh.Width = FrmMain.Width / 2 - 300
    CmdAbout.Width = FrmMain.Width / 2 - 300
    CmdAction.Width = FrmMain.Width / 2 - 300
    Drive1.Width = FrmMain.Width / 2 - 300
    CmdQuitter.Width = FrmMain.Width / 2 - 300
    File1.Width = FrmMain.Width / 2 - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub TxtChemin_Change()
    On Error Resume Next
    Dir1.Path = TxtChemin.Text
End Sub
