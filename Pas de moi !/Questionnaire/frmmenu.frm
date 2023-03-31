VERSION 5.00
Begin VB.Form frmmenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   -105
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   6
      Top             =   7800
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   960
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label opt3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "QUITTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label opt2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "A PROPOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4440
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label opt 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "JOUER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Copyright 2004        Application programmée par SDan"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   8640
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Le Questionnaire"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
opt.ForeColor = &HFF0000
opt2.ForeColor = &HFF0000
opt3.ForeColor = &HFF0000
End Sub
Private Sub opt_Click()
Load frmcharg
Unload frmmenu
End Sub
Private Sub opt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
opt.ForeColor = &HFFFF&
End Sub
Private Sub opt2_Click()
Load frmApp
End Sub
Private Sub opt2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
opt2.ForeColor = &HFFFF&
End Sub
Private Sub opt3_Click()
End
End Sub
Private Sub opt3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
opt3.ForeColor = &HFFFF&
End Sub

