VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF8080&
      Height          =   315
      ItemData        =   "frm1.frx":0000
      Left            =   0
      List            =   "frm1.frx":0010
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Choix"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1920
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "go"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Text1.Text = Time
End Sub

Private Sub Timer1_Timer()
If Combo1.Text = "Arrêter" Then
If Text1.Text = Text2.Text Then Shell "rundll32 shell32,SHExitWindowsEx 1", vbNormalFocus
End If
If Combo1.Text = "Redémarrer" Then
If Text1.Text = Text2.Text Then Shell "rundll32 shell32,SHExitWindowsEx 2", vbNormalFocus
End If
If Combo1.Text = "Déconnexion" Then
If Text1.Text = Text2.Text Then Shell "rundll32 shell32,SHExitWindowsEx 0", vbNormalFocus
End If
End Sub

Private Sub Timer2_Timer()
Text1.Text = Time
End Sub
