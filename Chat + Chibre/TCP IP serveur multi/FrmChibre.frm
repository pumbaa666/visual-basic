VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmChibre 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chibre"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkWsk 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   960
   End
   Begin MSWinsockLib.Winsock Wsk 
      Left            =   720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label LblTitre 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Le chibre d'la mort qui tue by Loïc Correvon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "(pauvre Loïc, ïc, ïc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "FrmChibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClkWsk_Timer()
    LblWsk.Caption = "State : " + Str(Wsk.State)
    LblRemote.Caption = "Remote : " + Str(Wsk.RemotePort)
    LblLocal.Caption = "Local : " + Str(Wsk.LocalPort)
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub FichierQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    Distribution
'    Call cdtInit(vLargeur, vHauteur)

'    Call cdtDrawExt(FrmMain.hdc, 150, 250, vLargeur * 2, vHauteur * 2, 0, &H0, vbBlue)
'    Call cdtDrawExt(FrmMain.hdc, 50, 150, vLargeur * 2, vHauteur * 2, 6, &H0, vbBlue)
'    Call cdtDrawExt(FrmMain.hdc, 250, 150, vLargeur * 2, vHauteur * 2, 45, &H0, vbBlue)
'    Call cdtDrawExt(FrmMain.hdc, 150, 50, vLargeur * 2, vHauteur * 2, 12, &H0, vbBlue)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vCarteCliquee As Double
Dim vCarte As Integer
Dim vAffCarte As Integer

    If Y > 5200 And Y < 6600 And X > 750 And X < 6015 And Button = 1 Then
        vCarteCliquee = Int((X - 750) / Int(35 * Screen.TwipsPerPixelX))
        If vCarteCliquee > 8 Then vCarteCliquee = 8

        CarteSel (vCarteCliquee)
        If vCarteEnCours(0) <> 20 And vCarteEnCours(0) <> 0 Then
            FrmMain.Cls
            vCouleur = tJeu(vCarteCliquee, 1)

            If vCarteEnCours(0) = 14 Then    ' Pour ce focking As !
                vAffCarte = vCarteEnCours(1) '(Merci Micro$oft :-@)
            Else
                vAffCarte = 4 * vCarteEnCours(0) - (4 - vCarteEnCours(1))
            End If

            Call cdtDrawExt(FrmMain.hdc, 150, 200, vLargeur, vHauteur, vAffCarte, &H0, vbBlue)
            Affichage
            FrmMain.Refresh
        End If
    End If
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    LblX.Caption = "X : " + Str(X) + " / " + Str(X / Screen.TwipsPerPixelX)
'    LblY.Caption = "Y : " + Str(Y) + " / " + Str(Y / Screen.TwipsPerPixelY)
'End Sub

Private Function CarteSel(ByVal vNumCarte As Integer)
    On Error Resume Next
    If tJeu(vNumCarte, 0) <> 21 Then
        vCarteEnCours(0) = tJeu(vNumCarte, 0)
        vCarteEnCours(1) = tJeu(vNumCarte, 1)
        tJeu(vNumCarte, 0) = 21
    ElseIf vNumCarte <> 0 And tJeu(vNumCarte - 1, 0) <> 21 Then
        vCarteEnCours(0) = tJeu(vNumCarte - 1, 0)
        vCarteEnCours(1) = tJeu(vNumCarte - 1, 1)
        tJeu(vNumCarte - 1, 0) = 21
    End If
End Function

Private Sub OptionHeberger_Click()
    Wsk.LocalPort = PortDistant
    Wsk.RemotePort = PortLocal
    Wsk.Listen
    FrmWait.Show
End Sub

Private Sub OptionRejoindre_Click()
    Wsk.LocalPort = PortLocal
    Wsk.RemotePort = PortDistant
    FrmIpServeur.Show
End Sub

Private Sub Wsk_DataArrival(ByVal bytesTotal As Long)
Dim vData As String

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

