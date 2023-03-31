VERSION 5.00
Begin VB.Form frm_param_map 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paramètres de la carte"
   ClientHeight    =   1455
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Musique"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   1045
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "Pas de fichier paramètre chargé"
         Top             =   250
         Width           =   3480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "App.path &&"
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   285
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nom de la carte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Text            =   "Pas de fichier paramètre chargé"
         Top             =   240
         Width           =   4445
      End
   End
   Begin VB.Menu Carte 
      Caption         =   "Carte"
      Index           =   1
      Begin VB.Menu Open 
         Caption         =   "Ouvrir..."
         Index           =   1
      End
      Begin VB.Menu save 
         Caption         =   "Sauver"
         Index           =   1
      End
   End
   Begin VB.Menu Mapedit 
      Caption         =   "MAPedit"
      Index           =   1
   End
   Begin VB.Menu quit 
      Caption         =   "Quitter"
      Index           =   1
   End
End
Attribute VB_Name = "frm_param_map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fic As String

Private Sub Open_Click(Index As Integer)
Dim dlg As CFileDialog
Set dlg = New CFileDialog
dlg.DialogTitle = "Choisissez un fond"
dlg.Filter = "Paramètres de cartes *.PRM|*.prm"
dlg.InitialDir = App.Path
If dlg.Show(False) Then
fic = dlg.FileName
Text1.Text = LireIni("Général", "Nom", fic)
Text2.Text = LireIni("Général", "musique", fic)
Else
End If

End Sub

Private Sub quit_Click(Index As Integer)
End
End Sub

Private Sub save_Click(Index As Integer)
Call EcrireIni("Général", "Nom", Text1.Text, fic)
If Check1.Value = 1 Then Call EcrireIni("Général", "musique", Text2.Text, fic) Else Call erireIni("Général", "musique", "", fic)
End Sub
