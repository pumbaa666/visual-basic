VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtRemote 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Text            =   "IP ou nom de l'hôte"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Timer ClkQuitter 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5520
      Top             =   5280
   End
   Begin MSWinsockLib.Winsock WSTCP 
      Left            =   4800
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1001
      LocalPort       =   1002
   End
   Begin VB.ListBox LstConnected 
      Height          =   2985
      ItemData        =   "FrmMain.frx":0000
      Left            =   4800
      List            =   "FrmMain.frx":000A
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox TxtNom 
      Height          =   375
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Votre nom ici"
      Top             =   120
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   10610
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"FrmMain.frx":0031
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "&Envoyer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox TxtSend 
      Enabled         =   0   'False
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   4455
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "&Connexion"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Shape ShpConnected 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label LblEtat 
      Caption         =   "Etat : Déconnecté"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vRemotePort As Integer

Private Sub ClkQuitter_Timer()
Static vCntQuitter As Boolean

    If vCntQuitter = False Then
        If WSTCP.State = 7 Then
            WSTCP.SendData ("<quit>" & "Admin: " & TxtNom.Text)
        End If
        vCntQuitter = True
    Else
        FermerPort
        End
    End If
End Sub

Private Sub CmdQuitter_Click()
    ClkQuitter.Enabled = True
End Sub

Private Sub Form_Load()
Dim vLocalPort As Integer
    Randomize
    vLocalPort = Int(Rnd * 400) + 1100
    WSTCP.Bind vLocalPort, WSTCP.LocalIP
    FermerPort
    FrmMain.Caption = "Chat " & WSTCP.LocalIP

    TxtRemote.Text = "172.16.118.14"

'    WSTCP.RemoteHost = "saule-loic"
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If Chr$(KeyAscii) = "<" Then
        MsgBox "Caractère interdit!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdConnect_Click
    End If
End Sub

Private Sub TxtRemote_KeyPress(KeyAscii As Integer)
    If Chr$(KeyAscii) = "<" Then
        MsgBox "Caractère interdit!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdConnect_Click
    End If
End Sub

Private Sub TxtSend_Change()
    If TxtSend.Text = "" Then
        CmdSend.Enabled = False
    Else
        CmdSend.Enabled = True
    End If
End Sub

Private Sub WSTCP_Close()
    MsgBox "L'hébérgeur s'est déconnecté", vbInformation, "Au revoir"
    LblEtat.Caption = "Etat : Déconnecté"
    ShpConnected.BackColor = &HFF&
    WSTCP.Close
    CmdConnect.Enabled = True
    TxtNom.Enabled = True
    TxtRemote.Enabled = True
'    End
End Sub

Private Sub WSTCP_Connect()
    WSTCP.SendData ("<new>" & TxtNom.Text)
    LblEtat.Caption = "Etat : Connecté"
    ShpConnected.BackColor = &HFF00&
    TxtSend.Enabled = True
    CmdConnect.Enabled = False
'    FrmTentative.Hide
End Sub

Private Sub WSTCP_DataArrival(ByVal bytesTotal As Long)
Dim vData As String
        
    WSTCP.GetData vData, vbString, bytesTotal
    If vData = "<clearliste>" Then
        LstConnected.Clear
        LstConnected.AddItem "Liste des personnes connectées"
        LstConnected.AddItem ""
    ElseIf Left(vData, 7) = "<liste>" Then
        LstConnected.AddItem Right(vData, Len(vData) - 7)
    ElseIf vData = "<kick>" Then
        MsgBox "Tu viens de te faire kické gros !", vbInformation, "Tcho"
    Else
        RTB.Text = RTB.Text & vData & Chr$(13)
'       RTB.SelStart = Len(RTB.Text) - 2
'        RTB.SelLength = 1
    End If
End Sub

Private Sub WSTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "L'erreur suivante c'est produite : " & Description, vbCritical, "Erreur"
'    If vRemotePort < 11 Then
'        vRemotePort = vRemotePort + 1
'        FermerPort
'        WSTCP.RemotePort = 1002 + vRemotePort
'        WSTCP.Connect
'        FrmTentative.LblTent.Caption = "Connexion au port n° 100" & vRemotePort
'        FrmTentative.ProgressBar.Value = FrmTentative.ProgressBar.Value + 10
'    Else
'        MsgBox "Impossible de se connecter", vbCritical, "Erreur"
'    End If
End Sub

Private Sub CmdConnect_Click()
    If Trim(LCase(TxtNom.Text)) = "votre nom ici" Or Trim(TxtNom.Text) = "" Then
        MsgBox "Veuillez choisir un pseudo", vbCritical, "Erreur"
    ElseIf Left(Trim(LCase(TxtNom.Text)), 5) = "admin" Then
        MsgBox "Non, non, non, ce pseudo est reservé", vbCritical, "Erreur"
    ElseIf TxtRemote.Text = "" Or TxtRemote.Text = "IP ou nom de l'hôte" Then
        MsgBox "Veuillez entrer l'IP ou le nom de l'hôte", vbCritical, "Erreur"
    Else
        'FrmTentative.Show
        WSTCP.RemoteHost = TxtRemote.Text
        TxtNom.Enabled = False
        TxtRemote.Enabled = False
        FermerPort
        WSTCP.Connect
    End If
End Sub

Private Sub CmdSend_Click()
    If WSTCP.State <> 7 And WSTCP.State <> 2 And WSTCP.State <> 1 And WSTCP.State <> 4 Then
        MsgBox "L'ordinateur distant est déconnecté, vous ne pouvez pas envoyer de messages.", vbCritical, "Erreur"
    ElseIf LstConnected.ListCount < 3 Then
        MsgBox "Veuillez patientez encore un instant svp", vbInformation, "Wait"
    Else
        If LCase(Trim(TxtSend.Text)) = "<cls>" Then
            RTB.Text = ""
        ElseIf LCase(Trim(TxtSend.Text)) = "<exit>" Then
            CmdQuitter_Click
        Else
            If Left(TxtSend.Text, 1) = "<" Then
                MsgBox "Opération incorrecte : " & TxtSend.Text, vbCritical, "Erreur"
            Else
                WSTCP.SendData TxtNom.Text & ": " & TxtSend.Text
'                RTB.Text = RTB.Text & TxtNom.Text & ": " & TxtSend.Text & Chr$(13)
'                RTB.SelStart = Len(RTB.Text) - 2
'                RTB.SelLength = 1
            End If
        End If
        TxtSend.Text = ""
        TxtSend.SetFocus
    End If
End Sub

Private Sub TxtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdSend_Click
    End If
End Sub

Function FermerPort()
    If WSTCP.State <> sckClosed Then
        WSTCP.Close
    End If
End Function
