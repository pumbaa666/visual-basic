VERSION 5.00
Begin VB.Form Scores 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFEECC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Les meilleurs scores"
   ClientHeight    =   4965
   ClientLeft      =   6015
   ClientTop       =   1440
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   4470
   Begin VB.PictureBox OK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   75
      Picture         =   "frmHTaquinSc.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   4335
      TabIndex        =   4
      Top             =   4500
      Width           =   4335
   End
   Begin VB.ComboBox NbrePieces 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox HighList 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   3450
      Left            =   75
      TabIndex        =   2
      Top             =   975
      Width           =   4335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      Caption         =   "    Durée       Date      Nom"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   75
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      Caption         =   "Carrés de         pièces"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   165
      Width           =   4335
   End
End
Attribute VB_Name = "Scores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    For i = 3 To 8
        NbrePieces.AddItem Str$(i)
    Next
    NbrePieces.ListIndex = Taquin.AHori.Value - 3
End Sub
Private Sub NbrePieces_Click()
    For i = 1 To HighList.ListCount
        HighList.RemoveItem 0
    Next i
    For i = 1 To 10
        HighList.AddItem tsHigh(NbrePieces.ListIndex + 1, i)
    Next i
End Sub
Private Sub NbrePieces_GotFocus()
    NbrePieces_Click
End Sub
Private Sub OK_Click()
    Scores.Hide
End Sub
Private Sub OK_GotFocus()
    NbrePieces_Click
End Sub

