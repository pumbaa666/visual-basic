VERSION 5.00
Begin VB.Form FrmExplorateur 
   Caption         =   "Explorateur"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox TxtName 
      Height          =   405
      Left            =   3000
      MaxLength       =   80
      TabIndex        =   3
      Text            =   "Entrez le nom du fichier"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CmdCreer 
      Caption         =   "&Créer le fichier ici"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmExplorateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    TxtName.Text = "Entrez le nom du fichier"
    FrmMain.Show
    FrmExplorateur.Hide
End Sub

Private Sub CmdCreer_Click()
    If TxtName.Text = "Entrez le nom du fichier" Or TxtName.Text = "" Then
        MsgBox "Veuillez entrer le nom du fichier", vbCritical, "Erreur"
    Else
        If Right(TxtName.Text, 4) <> ".txt" Then
            TxtName.Text = TxtName.Text & ".txt"
        End If
        Open Dir1.Path & "\" & TxtName.Text For Output As #1
        Close #1
        MsgBox "Le fichier est créé", vbInformation, "Création de fichier"
        CmdAnnuler_Click
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCreer_Click
    End If
End Sub
