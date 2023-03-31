VERSION 5.00
Begin VB.Form FrmIpServeur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adresse IP"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkConnect 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   600
      Top             =   1200
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox ComboIP 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Entrez ou sélectionnez l'adresse IP du serveur"
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "FrmIpServeur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tAllIP(50) As String

Private Sub ClkConnect_Timer()
Static vCntConnect As Boolean
Dim vCntTest As Integer
Dim vTemp As String

    If vCntConnect = False Then
        FrmOptMulti.Wsk.Connect ComboIP.Text, PortDistant
        vCntConnect = True
    Else
        On Error GoTo Erreur
        vTemp = FrmOptMulti.Wsk.State
        FrmOptMulti.Wsk.SendData "[CONNECTED]"
        MsgBox "Connexion établie", vbInformation, "OK"
        FrmOptMulti.Hide
        FrmChat.Show
        FrmMain.OptionChat.Enabled = True
        FrmMain.OptionCouleur.Enabled = True
    
        Open "c:\tron.ini" For Append As #1
        For vcnt = 0 To 50
            If tAllIP(vcnt) = ComboIP.Text Then
                Exit For
            End If
            If tAllIP(vcnt) = "" Then
                Print #1, ComboIP.Text
                Exit For
            End If
        Next
        Close #1
        FrmIpServeur.Hide
        Exit Sub
        ClkConnect.Enabled = False
    End If

Erreur:
    ClkConnect.Enabled = False
    MsgBox "Impossible de se connecter", vbCritical, "Erreur"
    FrmOptMulti.Wsk.Close
End Sub

Private Sub CmdAnnuler_Click()
    FrmIpServeur.Hide
End Sub

Private Sub CmdOk_Click()
    ClkConnect.Enabled = True
End Sub

Private Sub Form_Activate()
Dim vIp As String
Dim vCntIp As Integer

    On Error Resume Next
    Open "c:\tron.ini" For Input As #1
    While Not (EOF(1))
        Line Input #1, vIp
        If vCntIp <> 0 Then
            ComboIP.AddItem vIp
            tAllIP(vCntIp - 1) = vIp
        End If
        vCntIp = vCntIp + 1
    Wend
    Close #1
End Sub
