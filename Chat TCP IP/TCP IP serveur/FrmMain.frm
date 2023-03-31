VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP / IP"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock WSTCP 
      Left            =   3840
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox LstConnected 
      Height          =   3180
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Timer ClkConnect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4440
      Top             =   3720
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"FrmMain.frx":0000
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "&Envoyer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox TxtSend 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdQuitter_Click()
    If WSTCP.State <> sckClosed Then
        WSTCP.Close
    End If
    End
End Sub

Private Sub ClkConnect_Timer()
    If WSTCP.RemoteHostIP <> "" Then
        ClkConnect.Enabled = False
        TxtSend.Enabled = True
    End If
End Sub

Private Sub Form_Load()

    If WSTCP.State <> sckClosed Then
        WSTCP.Close
    End If

    WSTCP.Bind 1001, WSTCP.LocalIP

    If WSTCP.State <> sckClosed Then
        WSTCP.Close
    End If

    WSTCP.Listen
End Sub

Private Sub TxtSend_Change()
    If TxtSend.Text = "" Then
        CmdSend.Enabled = False
    Else
        CmdSend.Enabled = True
    End If
End Sub

'Private Sub WSTCP_Close(Index As Integer)
'    If WSTCP(0).State <> sckClosed Then
'        WSTCP(0).Close
'    End If
'End Sub

Private Sub WSTCP_Connect()
    TxtSend.Enabled = True
End Sub

Private Sub WSTCP_ConnectionRequest(ByVal requestID As Long)
    If WSTCP.State <> sckClosed Then
        WSTCP.Close
    End If
    WSTCP.Accept requestID
    LstConnected.AddItem WSTCP.RemoteHostIP

    ClkConnect.Enabled = True
End Sub

Private Sub WSTCP_DataArrival(ByVal bytesTotal As Long)
Dim vData As String
    WSTCP.GetData vData, vbString, bytesTotal
    If Left(vData, 3) <> "®½£" Then
        RTB.Text = RTB.Text & vData & Chr$(13)
    Else
        For vCount = 0 To (LstConnected.ListCount - 1)
            LstConnected.ListIndex = vCount
            If LstConnected.List(vCount) = Right(vData, Len(vData) - 3) Then
                LstConnected.RemoveItem vCount
                Exit For
            End If
        Next
    End If
End Sub

Private Sub WSTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "L'erreur suivante c'est produite : " & Description, vbCritical, "Erreur"
End Sub

Private Sub CmdSend_Click()
    If WSTCP.State <> 7 And WSTCP.State <> 2 And WSTCP.State <> 1 Then
        MsgBox "L'ordinateur distant est déconnecté, vous ne pouvez pas envoyer de messages.", vbCritical, "Erreur"
    Else
        If TxtSend.Text = "<cls>" Then
            RTB.Text = ""
        Else
            WSTCP.SendData TxtNom.Text & ": " & TxtSend.Text
            RTB.Text = RTB.Text & TxtNom.Text & ": " & TxtSend.Text & Chr$(13)
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
