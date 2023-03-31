VERSION 5.00
Begin VB.Form FrmExplorateur 
   Caption         =   "Explorateur"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3000
      Pattern         =   "*.bmp"
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Choisissez l'image à analyser"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmExplorateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    FrmMain.Show
    FrmExplorateur.Hide
End Sub

Private Sub CmdCreer_Click()
    If TxtName.Text = "Entrez le nom du fichier" Or TxtName.Text = "" Then
        MsgBox "Veuillez entrer le nom du fichier", vbCritical, "Erreur"
    Else
        Open Dir1.Path & "\" & TxtName.Text For Output As #1
        Close #1
        MsgBox "Le fichier est créé", vbInformation, "Création de fichier"
        CmdAnnuler_Click
    End If
End Sub

Private Sub CmdOk_Click()
    If File1.FileName = "" Then
        MsgBox "Veuillez choisire une image", vbCritical, "Erreur"
    Else
        FrmPixel.Show
        FrmExplorateur.Hide
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

Private Sub File1_DblClick()
    CmdOk_Click
End Sub
