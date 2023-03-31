VERSION 5.00
Begin VB.Form FrmStatistiques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistiques"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label LblReel 
      Caption         =   "Vitesse réele : "
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label LblBrute 
      Caption         =   "Vitesse brute :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label LblFautes 
      Caption         =   "Nombre de fautes : "
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label LblNbFrappes 
      Caption         =   "Nombre de frappes : "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "FrmStatistiques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    FrmStatistiques.Hide
End Sub
