VERSION 5.00
Begin VB.Form main 
   Caption         =   "Cube 3D (sans DirectX) par An-Mojeg"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Points"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Couleur"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WireFrame"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   3720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4065
      ScaleWidth      =   4185
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    mode = 1
End Sub

Private Sub Command2_Click()
    mode = 2
End Sub

Private Sub Command4_Click()
    mode = 0
End Sub
Private Sub Form_Load()
    init_cub
    draw_cub
    mode = 1
    Timer1.Interval = 100
End Sub
Private Sub Timer1_Timer()
    pic.Cls
    rotat_cub
    draw_cub
End Sub
