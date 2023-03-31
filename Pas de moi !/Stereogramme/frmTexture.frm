VERSION 5.00
Begin VB.Form frmTexture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Textures"
   ClientHeight    =   2820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox Textures 
      Height          =   1815
      ItemData        =   "frmTexture.frx":0000
      Left            =   120
      List            =   "frmTexture.frx":0002
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Textures:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image TextureImage 
      Height          =   2295
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim Fic As String
Dim Dossier As String

    'récupère la liste des fichiers textures
    Dossier = App.Path & "\Textures\*.bmp"
    Fic = Dir(Dossier)
    Textures.AddItem Fic
    While Fic <> ""
        Fic = Dir
        Textures.AddItem Fic
    Wend


    ' désactiver feuille principale
    frmOptions.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
     
     ' activer feuille principale
   frmOptions.Enabled = True

End Sub

Private Sub OKButton_Click()

If Textures.Text <> "" Then
    FichierNomLongTexture = App.Path & "\textures\" & Textures.Text
End If

MajTexture

Unload Me
End Sub

Private Sub CancelButton_Click()

Unload Me
End Sub

Private Sub Textures_Click()
Dim Fic As String

    Fic = App.Path & "\textures\" & Textures.Text
    If Fic <> App.Path & "\textures\" Then
        TextureImage.Picture = LoadPicture(Fic)
    End If
End Sub
