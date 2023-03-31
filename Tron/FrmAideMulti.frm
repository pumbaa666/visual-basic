VERSION 5.00
Begin VB.Form FrmAideMulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Initialiser une partie multi-joueurs"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNext 
      Caption         =   "Aide de &Jeu"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "L'hébergeur commencera la partie quand il le souhaite en pressant sur Start."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "Une fois la connexion établie, un message avertira les joueurs que c'est fait et une fenêtre de chat s'ouvrira."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "L'autre joueur devra choisir Rejoindre la partie et entrer l'adresse IP de l'hébergeur."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "L'hébergeur devra simplement cliquer sur Option/Partie/Multi-joueur/ et choisir Héberger la partie"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Pour cela, il faut qu'un des joueurs héberge la partie et l'autre qui le rejoingne."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Tron peut se jouer à 2."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAideMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNext_Click()
    FrmAideJeu.Show
    FrmAideMulti.Hide
End Sub

Private Sub CmdOk_Click()
    FrmAideMulti.Hide
End Sub
