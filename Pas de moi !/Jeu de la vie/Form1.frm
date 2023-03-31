VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                               UNE MUTE ! LA VIE"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Text            =   "3"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Text            =   "2"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Text            =   "3"
      Top             =   5280
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   14
      Top             =   5760
      Width           =   1095
      Begin VB.Label Label2 
         Caption         =   "Suffocation"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
      Begin VB.Label Label1 
         Caption         =   "Solitude"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
      Begin VB.Label Label2 
         Caption         =   "Naissance"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.CheckBox Option1 
      Caption         =   "ON"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "VIDE"
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "QUIT"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INFO"
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REGEN"
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox gfx 
      AutoRedraw      =   -1  'True
      DrawWidth       =   6
      Height          =   6060
      Index           =   1
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   5
      Top             =   0
      Width           =   6060
   End
   Begin VB.PictureBox gfx 
      AutoRedraw      =   -1  'True
      DrawWidth       =   6
      Height          =   6060
      Index           =   0
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
regenvie
End Sub

Private Sub Command10_Click()
vide
End Sub

Private Sub Command2_Click()
Form2.Text1.Text = " ----==***/ POURKOIA LA VIE \***==----" & vbCrLf & _
                      "   1er test de la vie : La Naissance" & vbCrLf & _
                      "   Une cellule mort devient vivante " & vbCrLf & _
                      "   si elle a exactement trois " & vbCrLf & _
                      "   cellules voisines vivantes" & vbCrLf & vbCrLf & _
                      "   2eme test de la vie : La Survie" & vbCrLf & _
                      "   Une cellule reste en vie tant qu'elle a deux" & vbCrLf & _
                      "   ou trois voisines vivantes" & vbCrLf & vbCrLf & _
                      "   3eme test de la vie : La Mort" & vbCrLf & _
                      "   Dans les autres cas la cellule meurt par " & vbCrLf & _
                      "   etouffement (+de 3 voisines vivantes )" & vbCrLf & _
                      "   ou de solitude ( - de 2 voisines )" & vbCrLf & _
                      "          Voila la vie donc ....." & vbCrLf & "             D41U5(c)2004"
                      


Form2.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Activate()
DoEvents
lavie
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub gfx_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
addcell X, Y
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = False Then
Option1.Caption = "ON"
Else
Option1.Caption = "OFF"
End If
End Sub

Private Sub Text1_Change(Index As Integer)
bmin = Val(Text1(0))
bmax = Val(Text1(1))
brev = Val(Text1(2))


End Sub
