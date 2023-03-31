VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajout score"
   ClientHeight    =   1260
   ClientLeft      =   195
   ClientTop       =   465
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxtAlderic 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox TxtLoic 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LblJoueur2 
      Caption         =   "Aldéric : "
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label LblJoueur1 
      Caption         =   "Loïc : "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()

    If IsNumeric(TxtLoic.Text) And IsNumeric(TxtAlderic.Text) Then
        FrmMain.LstAlderic.Clear
        FrmMain.LstLoic.Clear
        FrmMain.LstAlderic.AddItem TxtAlderic.Text
        FrmMain.LstLoic.AddItem TxtLoic.Text

        Open FrmFichier.TxtFichier.Text For Append As #1
        If Len(TxtLoic.Text) < 2 Then
            TxtLoic.Text = "0" & TxtLoic.Text
        End If

        If Len(TxtAlderic.Text) < 2 Then
            TxtAlderic.Text = "0" & TxtAlderic.Text
        End If

        Print #1, TxtLoic.Text & "-" & TxtAlderic.Text
        Close #1

        FrmAdd.Hide
        'Chargement
    Else
        MsgBox "Saisie incorrecte !", vbCritical, "Errors"
    End If
End Sub

Private Sub Form_Load()
    LblJoueur1.Caption = vJoueur1
    LblJoueur2.Caption = vJoueur2
End Sub

Private Sub TxtAlderic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtLoic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

