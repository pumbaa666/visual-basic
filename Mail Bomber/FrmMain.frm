VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail bomber"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtSujet 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "&Envoyer"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox TxtText 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtDest 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Sujet :"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Text :"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Destinataire :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdSend_Click()
Dim vCount As Integer
    For vCount = 0 To Int(TxtNombre.Text)
    Next
End Sub

Sub PhraseGenerator(ByVal vNbMots As Integer)

End Sub
