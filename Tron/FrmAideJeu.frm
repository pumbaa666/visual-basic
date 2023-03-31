VERSION 5.00
Begin VB.Form FrmAideJeu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comment jouer"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNext 
      Caption         =   "Aide &Scores"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   $"FrmAideJeu.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "En mode solo, vous pouvez mettre le jeu en pause en appuyant sur ENTRER"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Pour vous déplacer, utilisez les touches flèchées"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Le but du jeu est très simple : Survivre le plus longtemps possible et faire le plus grand score."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmAideJeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNext_Click()
    FrmAideScore.Show
    FrmAideJeu.Hide
End Sub

Private Sub CmdOk_Click()
    FrmAideJeu.Hide
End Sub
