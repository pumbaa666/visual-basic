VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PC shut down ! (version spécial farce)"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Form1.2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Cacher pour la farce"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   3000
      Picture         =   "Form1.2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "A&rrêter et redémarrer le systeme"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Arrêter le systeme"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   840
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444930
      CurrentDate     =   36829.5
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444930
      CurrentDate     =   36829.5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command3.Enabled = True
    DTP1.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    Command1.Enabled = False
    Timer2.Enabled = True
    Timer2.Interval = 1000
End Sub

Private Sub Command3_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
    DTP2.Value = Time
End Sub

Private Sub Timer1_Timer()
    DTP2.Value = Time
End Sub

Private Sub Timer2_Timer()
    If DTP1.Hour = DTP2.Hour And DTP1.Minute = DTP2.Minute And DTP1.Second = DTP2.Second And Option1.Value = True Then
        w95shutdown
        End
    ElseIf DTP1.Hour = DTP2.Hour And DTP1.Minute = DTP2.Minute And DTP1.Second = DTP2.Second And Option2.Value = True Then
        w95reboot
        End
    End If
End Sub
