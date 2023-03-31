VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmPeekA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connexion UDP"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRAZ 
      Caption         =   "Tout e&ffacer"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      TextRTF         =   $"FrmPeekA.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtPseudo 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Text            =   "Votre Pseudo"
      Top             =   840
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog DialFont 
      Left            =   5640
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdFont 
      Caption         =   "Couleur du &texte"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "&Envoyer"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox TxtSend 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox TxtDest 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "B107A-10SYS.3LITI.ch"
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Se &connecter à cet ordi"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock WskPeerA 
      Left            =   4320
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "FrmPeekA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vConnect As Boolean
Dim vNbCar As Double

Private Sub CmdConnect_Click()
    If vConnect = False Then
        With WskPeerA
            .RemoteHost = TxtDest.Text
            .RemotePort = 1002   ' Port auquel on se connecte.
            .Bind 1001           ' Établit le lien avec le port local.
        End With
        vConnect = True
    Else
        MsgBox "Vous êtes déjà connecté à un port.", vbCritical, "Erreur"
    End If
End Sub

Private Sub CmdFont_Click()
    DialFont.ShowColor
    TxtSend.ForeColor = DialFont.Color
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdRAZ_Click()
    RTB.Text = "Super Chat by Loïc" & Chr(13)
End Sub

Private Sub CmdSend_Click()
    If vConnect = True Then
        If TxtPseudo.Text = "Votre Pseudo" Then
            MsgBox "Veuillez indiquer votre pseudo.", vbCritical, "Erreur"
        Else
            TxtPseudo.Enabled = False
            If TxtSend.Text <> "" Then
                RTB.Text = RTB.Text & Chr(13) & TxtPseudo & " : " & TxtSend.Text
                RTB.Find TxtPseudo & " : " & TxtSend.Text, vNbCar
                vNbCar = Len(RTB.Text)
                RTB.SelColor = DialFont.Color
                WskPeerA.SendData TxtPseudo & " : " & TxtSend.Text
                TxtSend.Text = ""
                TxtSend.SetFocus
            End If
        End If
    Else
        MsgBox "Vous devez d'abord vous connecter avant d'envoyer des données.", vbCritical, "Erreur"
    End If
End Sub

Private Sub Form_Load()
    RTB.Text = "Super Chat by Loïc" & Chr(13)
    vNbCar = Len(RTB.Text)
End Sub

Private Sub TxtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdSend_Click
    End If
End Sub

Private Sub WskPeerA_DataArrival(ByVal bytesTotal As Long)
Dim vDataRecoit As String
    WskPeerA.GetData vDataRecoit
    RTB.Text = RTB.Text & Chr(13) & vDataRecoit
End Sub
