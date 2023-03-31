VERSION 5.00
Begin VB.Form frmcharg 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
   Begin VB.Timer bande 
      Interval        =   40
      Left            =   5400
      Top             =   5880
   End
   Begin VB.Timer ch1 
      Interval        =   300
      Left            =   9360
      Top             =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   6975
   End
   Begin VB.Label pour 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   735
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   6255
   End
   Begin VB.Label ch 
      BackColor       =   &H80000012&
      Caption         =   "Chargement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   4440
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Shape bnd 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   15
   End
   Begin VB.Label pour 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   2
      Top             =   4200
      Width           =   375
   End
End
Attribute VB_Name = "frmcharg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub term()
Unload q1
Unload frmIntroAfr
Load frmChoice
Unload frmcharg
End Sub
Private Sub bande_Timer()
pour(0).Caption = Val(bnd.Width / 6255 * 100)
If bnd.Width > 6255 Then bnd.Width = 6240
bnd.Width = bnd.Width + 40
If pour(0).Caption = 100 Then frmcharg.term
End Sub
Private Sub ch1_Timer()
If ch.Caption = "Chargement...." Then ch.Caption = "Chargement"
ch.Caption = ch.Caption & "."
End Sub

Private Sub Form_Load()
Show
End Sub
