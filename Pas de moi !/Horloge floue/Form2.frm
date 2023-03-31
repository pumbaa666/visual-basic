VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options horloge floue"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Affichage iconetray"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "Horloge floue"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Date actuelle"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Heure Exacte avec date"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Heure Exacte"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
      Begin VB.OptionButton Option1 
         Caption         =   "24 heures"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "12 heures"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Affichage du Texte sur le bureau"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   4
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Niveau du floue"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Mnu_Main"
      Visible         =   0   'False
      Begin VB.Menu Mnu_About 
         Caption         =   "A propos..."
      End
      Begin VB.Menu Mnu_Param 
         Caption         =   "Paramètres"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Cancel 
         Caption         =   "Annuler"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    SaveSetting App.Title, "Settings", "Desk", Str(Check1.Value)
End Sub

Private Sub Form_Load()
    HScroll1.Value = GetSetting(App.Title, "Settings", "Level", 0)  'Récupérartion du niveau de l'horloge
    HScroll1_Change 'Force HScoll à afficher le bon texte
    Option1(Val(GetSetting(App.Title, "Settings", "Format", "1"))).Value = True
    Option2(Val(GetSetting(App.Title, "Settings", "Systray", "3"))).Value = True
    Check1.Value = Val(GetSetting(App.Title, "Settings", "Desk", "1"))
End Sub

Private Sub HScroll1_Change()
    Select Case HScroll1.Value  'Selon l'emplacement du HScroll on détermine le texte à afficher
        Case 0
            Label1.Caption = "Niveau du floue : Précise"
        Case 1
            Label1.Caption = "Niveau du floue : Un peu"
        Case 2
            Label1.Caption = "Niveau du floue : Beaucoup"
        Case 3
            Label1.Caption = "Niveau du floue : A la folie"
        Case 4
            Label1.Caption = "Niveau du floue : Dans le vague"
    End Select
    SaveSetting App.Title, "Settings", "Level", HScroll1.Value  'Sauvegarde dans la base de registre du niveau de l'horloge
End Sub

Private Sub Mnu_About_Click()
    About.Show vbModal  'Mise en modal
End Sub

Private Sub Mnu_Cancel_Click()
    'J'ai placer ce commentaire pour pas que VB supprime ce sub car il ne contient rien
End Sub

Private Sub Mnu_Exit_Click()
    Unload Form1    'Ferme l'horloge    afin d'effacer l'icône dans le systray
    Unload Me   'Ferme cette form
    End 'Ferme le soft
End Sub

Private Sub Mnu_Param_Click()
    Form2.Show vbModal  'Affiche de cette fenêtre en modal
End Sub

Private Sub Option1_Click(Index As Integer)
    SaveSetting App.Title, "Settings", "Format", Index
End Sub

Private Sub Option2_Click(Index As Integer)
    SaveSetting App.Title, "Settings", "Systray", Str(Index)
End Sub
