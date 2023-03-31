VERSION 5.00
Begin VB.Form FrmPropos 
   Caption         =   "A propos..."
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "Web :"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label LblAR2 
      Caption         =   "http://membres.lycos.fr/pumbaa666"
      Height          =   255
      Left            =   1440
      MouseIcon       =   "FrmPropos.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Surveillance Port I/O Demo Version 1.0"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   2790
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright  © 2002, Loïc, Inc."
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   2070
   End
   Begin VB.Label LblMail 
      AutoSize        =   -1  'True
      Caption         =   "E-mail : pumbaa@net2000.ch"
      Height          =   195
      Left            =   960
      MouseIcon       =   "FrmPropos.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label LblBuffy 
      Caption         =   "http://membres.lycos.fr/buffyleguide "
      Height          =   195
      Left            =   1440
      MouseIcon       =   "FrmPropos.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1680
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmPropos.frx":091E
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "FrmPropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTemp As Integer

Private Sub CmdOk_Click()
    FrmMain.Show
    FrmPropos.Hide
End Sub

Private Sub LblAR2_Click()
    vTemp = Shell("c:\program files\internet explorer\IEXPLORE.EXE http://membres.lycos.fr/pumbaa666", vbMaximizedFocus)
End Sub

Private Sub LblBuffy_Click()
    vTemp = Shell("c:\program files\internet explorer\IEXPLORE.EXE http://membres.lycos.fr/buffyleguide", vbMaximizedFocus)
End Sub

Private Sub LblMail_Click()
    vTemp = Shell("c:\program files\internet explorer\IEXPLORE.EXE mailto: pumbaa@net2000.ch")
End Sub
