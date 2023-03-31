VERSION 5.00
Begin VB.Form FrmOuvrir 
   Caption         =   "Ouvrir"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmOuvrir.frx":0000
      Left            =   4800
      List            =   "FrmOuvrir.frx":0016
      TabIndex        =   7
      Text            =   "*.mp3"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CheckBox ChkKeep 
      Caption         =   "Ajouter à mes préférences"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3000
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Choisissez le fichier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmOuvrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    FrmMain.Show
    FrmOuvrir.Hide
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
Dim sOuvrir As StructListe
    If File1.FileName = "" Then
        MsgBox "Veuillez choisir un fichier", vbCritical, "Erreur"
    Else
        If ChkKeep.Value = Checked Then
            vNbPref = vNbPref + 1
            If Len(Dir1.Path) = 3 Then
                sOuvrir.vPath = Dir1.Path
            Else
                sOuvrir.vPath = Dir1.Path & "\"
            End If
            sOuvrir.vTitre = File1.FileName
            Open "c:\temp\prefmedia.dat" For Random As #1 Len = Len(sOuvrir)
            Put #1, vNbPref, sOuvrir
            Close #1
            FrmMain.ListPref.AddItem File1.FileName
        End If
        tMusique(0, vNbPref) = Trim(sOuvrir.vPath)
        tMusique(1, vNbPref) = Trim(sOuvrir.vTitre)
        FrmMain.LMP.Open Trim(sOuvrir.vPath) & Trim(sOuvrir.vTitre)
        FrmMain.Show
        FrmOuvrir.Hide
        vNumMus = vNbPref
    End If
End Sub

Private Sub Combo1_Click()
File1.Pattern = Combo1.Text
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
