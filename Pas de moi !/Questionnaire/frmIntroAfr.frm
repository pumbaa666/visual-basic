VERSION 5.00
Begin VB.Form frmIntroAfr 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pIC 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   1320
      ScaleHeight     =   1935
      ScaleWidth      =   2655
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton cmdno 
      Caption         =   "NON"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OUI"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Label lTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   5235
      End
   End
   Begin VB.Label lbli2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblq 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "frmIntroAfr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdno_Click()
Load frmChoice
Unload Me
End Sub
Public Sub cmdno_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Co = App.Path
If Right(Co, 1) <> "\" Then
Co = Co & "\"
End If
Co = Co & "Images\"
pIC.Picture = LoadPicture(Co & "f3.gif")
End Sub

Public Sub cmdok_Click()
Load q1
Unload frmIntroAfr
End Sub
Public Sub cmdok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C2 = App.Path
If Right(C2, 1) <> "\" Then
C2 = C2 & "\"
End If
C2 = C2 & "images\"
pIC.Picture = LoadPicture(C2 & "f2.gif")
End Sub

Public Sub Form_Load()
Show
lblq.Caption = "Souhaitez-vous continuer?"
C1 = App.Path
If Right(C1, 1) <> "\" Then
C1 = C1 & "\"
End If
C1 = C1 & "images\"
pIC.Picture = LoadPicture(C1 & "f1.gif")
Bon = 0
Mauv = 0
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1 = App.Path
If Right(C1, 1) <> "\" Then
C1 = C1 & "\"
End If
C1 = C1 & "images\"
pIC.Picture = LoadPicture(C1 & "f1.gif")
End Sub

