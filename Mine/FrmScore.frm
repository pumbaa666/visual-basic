VERSION 5.00
Begin VB.Form FrmScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Revenir au jeu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "&Effacer le score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label LblScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Meilleur score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "FrmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuitter_Click()
    FrmScore.Hide
    FrmJeu.Show
End Sub

Private Sub CmdReset_Click()
Dim vYes As Integer
    vYes = MsgBox("Etes-vous sur de vouloir effacer le score ?!?", vbYesNo)
    If vYes = vbYes Then
        Open "c:\temp\scoremine.dat" For Output As #1
        Print #1, "Anonyme 999"
        Close #1
    End If
    LblScore.Caption = "Anonyme 999"
End Sub

Private Sub Form_Load()
Dim vScore As String
Dim vTestFile As Boolean
    On Error GoTo CreatFile
    Open "c:\temp\scoremine.dat" For Input As #1
    vTestFile = 1
    Line Input #1, vScore
    Close #1
    LblScore.Caption = vScore & " secondes"
CreatFile:
    If vTestFile = 0 Then
        Open "c:\temp\scoremine.dat" For Append As #1
        Print #1, "Anonyme 999"
        Close #1
        LblScore.Caption = "Anonyme 999"
    End If
End Sub
