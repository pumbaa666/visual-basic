VERSION 5.00
Begin VB.Form FrmAnni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anniversaire"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMail 
      Caption         =   "&Envoyer un mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Timer ClkNom 
      Interval        =   250
      Left            =   240
      Top             =   1560
   End
   Begin VB.Label LblAge 
      Caption         =   "Et ça lui fait "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label LblNom 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Attention, aujourd'hui c'est l'anniversaire de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "FrmAnni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClkNom_Timer()
    LblNom.ForeColor = "&HFF" + Hex(Int(Rnd * 5000))
End Sub

Private Sub CmdMail_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & vMailAnni & "?subject=Anniversaire", vbMaximizedFocus
End Sub

Private Sub CmdQuitter_Click()
    FrmAnni.Hide
    FrmMain.Show
End Sub

Private Sub Form_Load()
Dim vAge As Integer
    LblNom.Caption = vNom
    vAge = Year(Date) - Year(CDate(vDate))
    LblAge.Caption = LblAge.Caption & vAge & " ans"
End Sub
