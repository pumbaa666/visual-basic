VERSION 5.00
Begin VB.Form FrmAPropos 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   240
   ClientTop       =   390
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label LblAR2 
      Caption         =   "http://membres.lycos.fr/pumbaa666"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAPropos.frx":0000
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label LblBuffy 
      Caption         =   "http://membres.lycos.fr/buffyleguide"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAPropos.frx":0442
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   3600
      Picture         =   "FrmAPropos.frx":0884
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label Label3 
      Caption         =   "Ce programme est totalement libre d'utilisation."
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label LblMail 
      Caption         =   "pumbaa@net2000.ch"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAPropos.frx":1BA3
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Développé par Loïc Correvon"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label LblMM 
      Caption         =   "http://membres.lycos.fr/manson666marilyn"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAPropos.frx":1FE5
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
End
Attribute VB_Name = "FrmAPropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    Hide
End Sub
