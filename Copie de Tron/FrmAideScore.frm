VERSION 5.00
Begin VB.Form FrmAideScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Les scores"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNext 
      Caption         =   "Aide &Difficulté"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "fois le nombre de changement de direction divisé par 10"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "fois 2 si vous avez choisi les 2 modes ci-dessus"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "fois 3 si vous avez choisi le mode ""aléatoire"""
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "fois 2 si vous avez mis des obstacles"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "fois la difficulté choisie (Facile = 1; Moyen = 2; Difficile = 3)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Le nombre de secondes que vous avez survécu fois 100"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Les scores sont calculés de cette manière : "
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "FrmAideScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNext_Click()
    FrmAideDif.Show
    FrmAideScore.Hide
End Sub

Private Sub CmdOk_Click()
    FrmAideScore.Hide
End Sub
