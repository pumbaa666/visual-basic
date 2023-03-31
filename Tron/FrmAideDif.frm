VERSION 5.00
Begin VB.Form FrmAideDif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Les difficultés"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNext 
      Caption         =   "Aide &Partie Multi"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Il changera plus ou moins souvent selon le degré de difficulté choisi."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Fais subitement changer de direction votre Tron."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Le mode aléatoire : "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Ajoute 50, 100 ou 150 obstacles sur l'aire de jeu (proportionel à la difficulté.)"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Les obstacles : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAideDif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNext_Click()
    FrmAideMulti.Show
    FrmAideDif.Hide
End Sub

Private Sub CmdOk_Click()
    FrmAideDif.Hide
End Sub
