VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmClient2 
   Caption         =   "Client 2"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkConnect 
      Interval        =   50
      Left            =   240
      Top             =   840
   End
   Begin MSWinsockLib.Winsock WSClient 
      Left            =   960
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmClient2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vDonnees As String

Private Sub ClkConnect_Timer()
    If WSClient.RemoteHostIP <> "" And vNumPort < 10 Then
        ClkConnect.Enabled = False
        vNumPort = vNumPort + 1
    End If
End Sub

Private Sub Form_Load()
    FermerPort
    WSClient.Bind vNumPort, WSClient.LocalIP
    FermerPort
    WSClient.Listen
End Sub

Function FermerPort()
    If WSClient.State <> sckClosed Then
        WSClient.Close
    End If
End Function

Private Sub WSClient_ConnectionRequest(ByVal requestID As Long)
    FermerPort
    WSClient.Accept requestID
    FrmMain.LstConnected.AddItem WSClient.RemoteHostIP

    ClkConnect.Enabled = True
End Sub

Private Sub WSClient_DataArrival(ByVal bytesTotal As Long)
Dim vData As String
Dim vCount As Integer
    WSClient.GetData vData, vbString, bytesTotal
    vDonnees = vData
    If Left(vData, 6) <> "<quit>" Then
        FrmMain.RTB.Text = FrmMain.RTB.Text & vData & Chr$(13)
        FrmMain.RTB.SelStart = Len(FrmMain.RTB.Text) - 2
        FrmMain.RTB.SelLength = 1

        WSClient.SendData vData
    Else
        For vCount = 0 To (FrmMain.LstConnected.ListCount - 1)
            FrmMain.LstConnected.ListIndex = vCount
            If FrmMain.LstConnected.List(vCount) = Right(vData, Len(vData) - 6) Then
                FrmMain.LstConnected.RemoveItem vCount
                Exit For
            End If
        Next
    End If
End Sub

Private Sub WSClient_SendComplete()
Dim vNewData As String
    If FrmMain.WSTCP.RemoteHostIP <> "" And vDonnees <> "" Then
        vNewData = vDonnees
        vDonnees = ""
        FrmMain.WSTCP.SendData vNewData
    End If
End Sub
