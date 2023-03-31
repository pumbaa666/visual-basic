VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Choix d'une direction"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2835
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   1785
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btdir 
      Caption         =   "Nord"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Btdir 
      Caption         =   "Ouest"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Btdir 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Btdir 
      Caption         =   "Sud"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Btdir 
      Caption         =   "Est"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btdir_Click(Index As Integer)
Dim I As Long

For I = 0 To UBound(Pas, 2)
    If Form2.Bloc(I).BackColor = 16773103 And Left$(Form2.Bloc(I).Caption, 2) = "DD" Then
        Form2.Bloc(I).Caption = "DD_" & Btdir(Index).Index
    End If
Next I

Unload Me
End Sub
