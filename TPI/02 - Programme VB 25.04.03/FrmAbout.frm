VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Dice Value"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Projet de TPI Informaticien Visual Basic"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Créé par Loïc Correvon"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3360
      Picture         =   "FrmAbout.frx":0000
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    FrmMain.Show
    FrmAbout.Hide
End Sub

