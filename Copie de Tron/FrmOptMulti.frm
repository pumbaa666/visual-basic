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
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer ClkConnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
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
Private Sub ClkConnect_Timer()
Dim vIp As String
Static vCntConnect As Boolean

    If vCntConnect = False Then
        Wsk.LocalPort = PortLocal
        vIp = InputBox("Entrez l'adresse IP du serveur", "IP Distante")
        Wsk.Connect vIp, PortDistant
        vCntConnect = True
    Else
        On Error GoTo Erreur
        Wsk.SendData "[CONNECTED]"
        MsgBox "Connexion établie", vbInformation, "OK"
        ClkConnect.Enabled = False
        FrmOptMulti.Hide
        FrmChat.Show
        FrmMain.OptionChat.Enabled = True
        FrmMain.ShpTerrain(0).FillColor = vRemoteCouleur
        FrmMain.OptionCouleur.Enabled = True
    End If
    Exit Sub
Erreur:
    MsgBox "Impossible de se connecter", vbCritical, "Erreur"
    ClkConnect.Enabled = False
    Wsk.Close
End Sub

Private Sub CmdClient_Click()
    tCoo(X) = 59
    tCoo(Y) = 35
    vDX = -1
    vDY = 0
    FrmMain.ShpTerrain(2159).FillColor = vCouleur

    ClkConnect.Enabled = True

    vQui = "Client"
    FrmMain.CmdMulti.Enabled = False
End Sub

Private Sub CmdServeur_Click()
    tCoo(X) = 0
    tCoo(Y) = 0
    vDX = 1
    vDY = 0
    If vCouleur = "" Then vCouleur = "&H0"
    FrmMain.ShpTerrain(0).FillColor = vCouleur

    Wsk.LocalPort = PortDistant
    Wsk.Listen
    FrmOptMulti.Hide
    FrmWait.Show

    vQui = "Serveur"
    FrmMain.CmdMulti.Enabled = False
End Sub

Private Sub Form_Load()
    vRemoteCouleur = "&H0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.CmdMulti.Visible = True
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
        FrmChat.Show
        FrmMain.OptionCouleur.Enabled = True
        FrmMain.ShpTerrain(2159).FillColor = vCouleur
        FrmMain.OptionCouleur.Enabled = True
        FrmMain.CmdStart.Enabled = True

    ElseIf vData = "[START]" Then
        FrmMain.Picture1.Visible = True
        FrmMain.Picture1.SetFocus
        FrmMain.ClkMain.Enabled = True
        ClearTerrain
        Start

    ElseIf Left(vData, 9) = "[COULEUR]" Then
        vRemoteCouleur = Right(vData, Len(vData) - 9)
        If vQui = "Client" Then
            FrmMain.ShpTerrain(0).FillColor = vRemoteCouleur
        Else
            FrmMain.ShpTerrain(2159).FillColor = vRemoteCouleur
        End If

    ElseIf Left(vData, 5) = "[COO]" Then
        tRemoteCoo(X) = Int(Mid(vData, 6, 2))
        tRemoteCoo(Y) = Int(Mid(vData, 9, 2))
        tTerrain(tRemoteCoo(X), tRemoteCoo(Y)) = 1
        FrmMain.ShpTerrain(tRemoteCoo(Y) * 60 + tRemoteCoo(X)).FillColor = vRemoteCouleur

    ElseIf vData = "[PERDU]" Then
        FrmMain.ClkMain.Enabled = False
        FrmMain.CmdStart.Enabled = True
        MsgBox "Vous avez gagné", vbInformation, "Gagné"

    ElseIf vData = "[QUIT]" Then
        FrmMain.ClkMain.Enabled = False
        Wsk.Close
        FrmMain.Picture1.Visible = False
        MsgBox "Votre adversaire c'est déconnecté", vbInformation, "Déconnexion"

    Else
        FrmChat.RTB.Text = FrmChat.RTB.Text & vData & Chr$(13)
    End If
End Sub
