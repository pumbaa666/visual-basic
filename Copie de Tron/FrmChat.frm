VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdHide 
      Caption         =   "&Masquer le chat"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton CmdPseudo 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox TxtSend 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox TxtPseudo 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Text            =   "Votre Pseudo"
      Top             =   120
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
      _Version        =   393217
      TextRTF         =   $"FrmChat.frx":0000
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "&Envoyer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Label LblRemoteIP 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label LblLocalIP 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label LblPseudo 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdHide_Click()
    FrmChat.Hide
End Sub

Private Sub CmdPseudo_Click()
    If TxtPseudo = "" Or TxtPseudo = "Votre Pseudo" Then
        MsgBox "Veuillez entrer un pseudo", vbCritical, "Erreur"
    Else
        LblPseudo.Caption = "Votre Pseudo : " & TxtPseudo.Text
        TxtPseudo.Visible = False
        CmdPseudo.Visible = False
        CmdSend.Enabled = True
    End If
End Sub

Private Sub CmdSend_Click()
    If FrmOptMulti.Wsk.State <> 7 Then
        MsgBox "L'ordinateur distant est déconnecté, vous ne pouvez pas envoyer de messages.", vbCritical, "Erreur"
    ElseIf TxtPseudo.Visible = True Then
        MsgBox "Entrez votre pseudo", vbCritical, "Erreur"
    Else
        If TxtSend.Text = "<cls>" Then
            RTB.Text = ""
        ElseIf TxtSend.Text <> "" Then
            FrmOptMulti.Wsk.SendData TxtPseudo.Text & ": " & TxtSend.Text
            RTB.Text = RTB.Text & TxtPseudo.Text & ": " & TxtSend.Text & Chr$(13)
        End If
        TxtSend.Text = ""
        TxtSend.SetFocus
    End If
End Sub

Private Sub Form_Load()
    LblLocalIP.Caption = "Votre adresse IP : " & FrmOptMulti.Wsk.LocalIP
    LblRemoteIP.Caption = "Son adresse IP : " & FrmOptMulti.Wsk.RemoteHostIP
End Sub

Private Sub TxtPseudo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdPseudo_Click
    End If
End Sub

Private Sub TxtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdSend_Click
    End If
End Sub
