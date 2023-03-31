VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmOptMulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choix"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2220
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Wsk 
      Left            =   360
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdClient 
      Caption         =   "&Rejoindre une partie (client)"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton CmdServeur 
      Caption         =   "&Hébérger une partie (serveur)"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmOptMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClient_Click()
    Wsk.LocalPort = PortLocal
    FrmIpServeur.Show
End Sub

Private Sub CmdServeur_Click()
    Wsk.LocalPort = PortDistant
    Wsk.Listen
    FrmOptMulti.Hide
    FrmWait.Show

    vQui = "Serveur"
End Sub

Private Sub Wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "L'erreur suivante c'est produite : " & Description, vbCritical, "Erreur"
End Sub

Private Sub Wsk_ConnectionRequest(ByVal requestID As Long)
    Wsk.Close
    Wsk.Accept requestID
End Sub

Private Sub Wsk_DataArrival(ByVal bytesTotal As Long)
Dim vData As String
Dim tRemoteCoo(1) As Integer

    Wsk.GetData vData, vbString, bytesTotal
    If vData = "[CONNECTED]" Then
        FrmWait.ClkProgress.Enabled = False
        FrmWait.Hide
        FrmOptMulti.Hide
        MsgBox "Connexion établie", vbInformation, "OK"
'        FrmChat.Show

    Else
'        FrmChat.RTB.Text = FrmChat.RTB.Text & vData & Chr$(13)
    End If
End Sub
