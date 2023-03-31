VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label LblMM 
      Caption         =   "http://membres.lycos.fr/manson666marilyn"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":0000
      MousePointer    =   2  'Cross
      TabIndex        =   8
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Développé par Loïc Correvon"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label LblMail 
      Caption         =   "pumbaa@net2000.ch"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":0442
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Ce programme est totalement libre d'utilisation."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Passez-le à vos amis ;-)"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   3600
      Picture         =   "FrmAbout.frx":0884
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "Version 1.1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label LblBuffy 
      Caption         =   "http://membres.lycos.fr/buffyleguide"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":3686
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label LblAR2 
      Caption         =   "http://membres.lycos.fr/pumbaa666"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":3AC8
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
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

Private Sub LblAR2_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & LblAR2.Caption, vbMaximizedFocus
End Sub

Private Sub LblBuffy_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & LblBuffy.Caption, vbMaximizedFocus
End Sub

Private Sub LblMail_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & LblMail.Caption, vbMaximizedFocus
End Sub

Private Sub LblMM_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & LblMM.Caption, vbMaximizedFocus
End Sub
