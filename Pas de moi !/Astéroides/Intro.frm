VERSION 5.00
Begin VB.Form frmIntro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   Picture         =   "Intro.frx":0000
   ScaleHeight     =   3315
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1680
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
frmAste.Show
Unload frmIntro
End Sub
