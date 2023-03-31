VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Ascii"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdTable 
      Caption         =   "&Voir la table"
      Height          =   1335
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdBizarre 
      Caption         =   "&Bizarre"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton CmdConv2 
      Caption         =   "Convertir"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton CmdConv 
      Caption         =   "Convertir"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TxtChar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Si la touche dont vous voulez connaitre le code ASCII ne correspond à aucun caractère, cliquer sur Bizarre et taper cette touche."
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Entrez le code dont vous voulez connaitre le caractère"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Entrez le caractère dont vous voulez connaitre le code ASCII"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBizarre_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox KeyCode
End Sub

Private Sub CmdConv_Click()
    If TxtChar.Text = "" Then
        MsgBox "Il n'y a rien a convertir... CONNARD!!!", vbCritical, "Troudbit"
    Else
        Label.Caption = "Le caractère " & TxtChar.Text & " vaut " & Asc(TxtChar.Text)
    End If
End Sub

Private Sub CmdConv2_Click()
    If TxtCode.Text = "" Then
        MsgBox "Il n'y a rien a convertir... CONNARD!!!", vbCritical, "Troudbit"
    Else
        Label2.Caption = "Le code " & TxtCode.Text & " correspond à " & Chr(TxtCode.Text)
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdTable_Click()
    FrmTable.Show
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdConv2_Click
    End If
End Sub

Private Sub TxtChar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdConv_Click
    End If
End Sub

