VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   Caption         =   "Anonyme Mail et Main Bomber"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   18
      Text            =   "1"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Status:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   3255
      Begin VB.Label StatusTxt 
         Caption         =   "http://www.timp3.t2u.com. By LJP "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox txtEmailServer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "mail.net2000.ch"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox ToNametxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtFromName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtEmailBodyOfMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtEmailSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtToEmailAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtFromEmailAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "&Envoyer"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6840
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label83 
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "#:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Serveur"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Son nom:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Ton nom:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Subject"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "To (e-mail):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "De (e-mail):"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim start As Single, Tmr As Single



Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.Connect ' Start connection
    
    WaitFor ("220")
    
    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    StatusTxt.Caption = "Connected"
    StatusTxt.Refresh

    Winsock1.SendData (first)

    StatusTxt.Caption = "Sending Message"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    StatusTxt.Caption = "Disconnecting"
    StatusTxt.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = ""
End Sub


Private Sub CmdSend_Click()
b = 0
Do
    b = b + 1
    Label83.Caption = "NB d'envoi effectués:" & b

    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text

    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    Close

    If Text1.Text = b Then
        Exit Do
        Label83.Caption = "yes"
    End If
Loop '
End Sub

Private Sub CmdQuitter_Click()
    Unload Me
    End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response

End Sub
