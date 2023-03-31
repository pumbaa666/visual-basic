VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Web"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
      Begin VB.Label LblAR2 
         Caption         =   "http://membres.lycos.fr/pumbaa666"
         Height          =   255
         Left            =   120
         MouseIcon       =   "FrmAbout.frx":0000
         MousePointer    =   2  'Cross
         TabIndex        =   10
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label LblBuffy 
         Caption         =   "http://membres.lycos.fr/buffyleguide"
         Height          =   255
         Left            =   120
         MouseIcon       =   "FrmAbout.frx":0442
         MousePointer    =   2  'Cross
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label LblMM 
         Caption         =   "http://membres.lycos.fr/manson666marilyn"
         Height          =   255
         Left            =   120
         MouseIcon       =   "FrmAbout.frx":0884
         MousePointer    =   2  'Cross
         TabIndex        =   8
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mail"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3495
      Begin VB.Label LblMail 
         Caption         =   "pumbaa@net2000.ch"
         Height          =   255
         Left            =   120
         MouseIcon       =   "FrmAbout.frx":0CC6
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   3840
      Picture         =   "FrmAbout.frx":1108
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Passez-le à vos amis ;-)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Ce programme est totalement libre d'utilisation."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Développé par Loïc Correvon"
      Height          =   255
      Left            =   120
      TabIndex        =   1
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
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & LblBuffy.Caption, vbMaximizedFocus
End Sub

Private Sub LblMail_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & LblMail.Caption, vbMaximizedFocus
End Sub

Private Sub LblMM_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & LblMM.Caption, vbMaximizedFocus
End Sub

