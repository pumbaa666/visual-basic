VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos..."
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      Caption         =   "Version 1.07"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Passez-le à vos amis ;-)"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Ce programme est totalement libre d'utilisation."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label LblMail 
      Caption         =   "pumbaa@net2000.ch"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2400
      MouseIcon       =   "FrmAbout.frx":0000
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Développé par Loïc Correvon : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
