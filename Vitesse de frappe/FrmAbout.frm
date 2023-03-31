VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos de SpeedHit"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAbout.frx":0000
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label LblAR2 
      Caption         =   "http://membres.lycos.fr/pumbaa666"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":005A
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label LblBuffy 
      Caption         =   "http://membres.lycos.fr/buffyleguide"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":049C
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Version 1.1"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   3720
      Picture         =   "FrmAbout.frx":08DE
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Passez-le à vos amis ;-)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Ce programme est totalement libre d'utilisation."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label LblMail 
      Caption         =   "pumbaa@net2000.ch"
      Height          =   255
      Left            =   240
      MouseIcon       =   "FrmAbout.frx":1BE0
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Développé par Loïc Correvon"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
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
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & LblBuffy.Caption
End Sub

Private Sub LblMail_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & LblMail.Caption, vbMaximizedFocus
End Sub
