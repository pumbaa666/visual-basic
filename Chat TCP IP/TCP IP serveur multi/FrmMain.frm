VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serveur"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkKick 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6840
      Top             =   6000
   End
   Begin VB.Timer ClkConnect 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5400
      Top             =   6000
   End
   Begin VB.Timer ClkList 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   6360
      Top             =   6000
   End
   Begin VB.Timer ClkEnvoie 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   6000
   End
   Begin MSWinsockLib.Winsock WSTCP 
      Index           =   0
      Left            =   4800
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox LstConnected 
      Height          =   4545
      ItemData        =   "FrmMain.frx":0000
      Left            =   4800
      List            =   "FrmMain.frx":000A
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   6015
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   1
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox TxtSend 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   4455
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label LblIp 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNbConnect As Integer
Dim vDeconnect As Integer
Dim vAffichage As String
Dim vKick As Integer

Private Sub ClkEnvoie_Timer()
Static vCount As Integer
Dim vTest As Boolean

    If WSTCP(vCount).State <> 7 And WSTCP(vCount).State <> 2 And WSTCP(vCount).State <> 1 Then
'        MsgBox "L'ordinateur distant est déconnecté, vous ne pouvez pas envoyer de messages.", vbCritical, "Erreur"
    Else
        If LCase(Trim(TxtSend.Text)) = "<cls>" Then
            RTB.Text = ""
        ElseIf LCase(Trim(TxtSend.Text)) = "<exit>" Then
            CmdQuitter_Click
        Else
            If WSTCP(vCount).State = 7 Then
                WSTCP(vCount).SendData vAffichage
            End If

            If vCount = 0 Then
                RTB.Text = RTB.Text & vAffichage & Chr$(13)
            End If
        End If
    End If
    If Left(vData, 6) = "<quit>" Then
        FermerPort vCount
    End If
    vCount = vCount + 1
    If vCount >= vNbConnect Then
        ClkEnvoie.Enabled = False
        TxtSend.Text = ""
        TxtSend.SetFocus
        vCount = 0
    End If
End Sub

Private Sub ClkKick_Timer()
Static vCntKick As Boolean

    If vCntKick = False Then
        WSTCP(vKick).SendData "<kick>"
        vCntKick = True
    Else
        WSTCP(vKick).Close
        ClkKick.Enabled = False
    End If
End Sub

Private Sub ClkList_Timer()
Static vCntListe As Integer
Static vCntUser As Integer

    If vCntListe = 0 And WSTCP(vCntUser).State = 7 Then
        WSTCP(vCntUser).SendData "<clearliste>"
    ElseIf WSTCP(vCntUser).State = 7 Then
        LstConnected.ListIndex = vCntListe + 1 - vDeconnect
        WSTCP(vCntUser).SendData "<liste>" & LstConnected.Text
    End If
    vCntListe = vCntListe + 1
    If vCntListe >= vNbConnect Then
        vCntUser = vCntUser + 1
        vCntListe = 0
        If vCntUser >= vNbConnect Then
            ClkList.Enabled = False
            vCntUser = 0
        End If
    End If
End Sub

Private Sub CmdQuitter_Click()
    FermerPort 0
    End
End Sub

Private Sub ClkConnect_Timer()
    If WSTCP(vNbConnect).RemoteHostIP <> "" And vNumPort < 11 Then
        ClkConnect.Enabled = False
        TxtSend.Enabled = True
        vNumPort = vNumPort + 1
        vNbConnect = vNbConnect + 1

        Load WSTCP(vNbConnect)
        FermerPort vNbConnect
        WSTCP(vNbConnect).Bind 1001, WSTCP(vNbConnect).LocalIP
        FermerPort vNbConnect
        WSTCP(vNbConnect).Listen
        vNumPort = 1002
    End If
End Sub

Private Sub Form_Load()
    FermerPort vNbConnect
    WSTCP(vNbConnect).Bind 1001, WSTCP(vNbConnect).LocalIP
    FermerPort vNbConnect
    WSTCP(vNbConnect).Listen
    vNumPort = 1002
    LblIp.Caption = "IP : " & WSTCP(0).LocalIP
End Sub

Private Sub LstConnected_KeyDown(KeyCode As Integer, Shift As Integer)
    If LstConnected.ListIndex > 1 And KeyCode = 46 Then
        vKick = LstConnected.ListIndex - 2
        ClkKick.Enabled = True
        vAffichage = "Admin: " & LstConnected.Text & " s'est fait bannir"
        ClkEnvoie.Enabled = True
        LstConnected.RemoveItem LstConnected.ListIndex
        vDeconnect = vDeconnect + 1
        ClkList.Enabled = True
    End If
End Sub

Private Sub TxtSend_Change()
    If TxtSend.Text = "" Then
        CmdSend.Enabled = False
    Else
        CmdSend.Enabled = True
    End If
End Sub

Private Sub WSTCP_Close(Index As Integer)
    FermerPort Index
End Sub

Private Sub WSTCP_Connect(Index As Integer)
    TxtSend.Enabled = True
End Sub

Private Sub WSTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    FermerPort Index
    WSTCP(Index).Accept requestID
    ClkConnect.Enabled = True
End Sub

Private Sub WSTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim vData As String
Dim vCount As Integer

    WSTCP(Index).GetData vData, vbString, bytesTotal
    If Left(vData, 6) = "<quit>" Then
        vAffichage = Right(vData, Len(vData) - 6) & " nous à quitté, paix à son âme"
        ClkEnvoie.Enabled = True
        ClkList.Enabled = True

        For vCount = 2 To vNbConnect + 1 - vDeconnect
            LstConnected.ListIndex = vCount
            If LstConnected.List(vCount) = Right(vData, Len(vData) - 13) Then
                LstConnected.RemoveItem vCount
                Exit For
            End If
        Next
        vDeconnect = vDeconnect + 1
    ElseIf Left(vData, 5) = "<new>" Then
        LstConnected.AddItem Right(vData, Len(vData) - 5)
        vAffichage = "Admin: " & Right(vData, Len(vData) - 5) & " nous a rejoint"
        ClkEnvoie.Enabled = True
        ClkList.Enabled = True
    Else
        vAffichage = vData
        ClkEnvoie.Enabled = True
    End If
End Sub

Private Sub WSTCP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "L'erreur suivante c'est produite : " & Description, vbCritical, "Erreur"
End Sub

Private Sub CmdSend_Click()
    vAffichage = "Admin: " & TxtSend.Text
    ClkEnvoie.Enabled = True
End Sub

Private Sub TxtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtSend.Text) <> "" Then
        CmdSend_Click
    End If
End Sub

Function FermerPort(ByVal vNumFermer As Integer)
    If WSTCP(vNumFermer).State <> sckClosed Then
        WSTCP(vNumFermer).Close
    End If
End Function
