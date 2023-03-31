VERSION 5.00
Begin VB.Form FrmAbout 
   Caption         =   "About"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmAbout.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LblBuffy 
      Caption         =   "http://membres.lycos.fr/buffyleguide "
      Height          =   195
      Left            =   1440
      MouseIcon       =   "FrmAbout.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1440
      Width           =   2745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "E-mail : pumbaa@net2000.ch"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   2100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright  © 2003, Loïc, Inc."
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gestion de contact Version 1.0"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   2190
   End
   Begin VB.Label LblAR2 
      Caption         =   "http://membres.lycos.fr/pumbaa666"
      Height          =   255
      Left            =   1440
      MouseIcon       =   "FrmAbout.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Web :"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTemp As Integer

Private Sub CmdOk_Click()
    FrmMain.Show
    FrmAbout.Hide
End Sub

Private Sub Label3_Click()
    vTemp = Shell("c:\program files\internet explorer\IEXPLORE.EXE mailto: pumbaa@net2000.ch", vbMaximizedFocus)
End Sub

Private Sub LblAR2_Click()
    vTemp = Shell("c:\program files\internet explorer\IEXPLORE.EXE http://membres.lycos.fr/pumbaa666", vbMaximizedFocus)
End Sub

Private Sub LblBuffy_Click()
    vTemp = Shell("c:\program files\internet explorer\IEXPLORE.EXE http://membres.lycos.fr/buffyleguide", vbMaximizedFocus)
End Sub

