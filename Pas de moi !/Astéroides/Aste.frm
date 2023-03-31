VERSION 5.00
Begin VB.Form frmAste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Astéroïds!"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   Icon            =   "Aste.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "Aste.frx":0E42
   ScaleHeight     =   7950
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerLevel 
      Interval        =   10
      Left            =   9240
      Top             =   3480
   End
   Begin VB.Timer TV0 
      Interval        =   40
      Left            =   9240
      Top             =   5880
   End
   Begin VB.Timer TimerAst 
      Interval        =   30
      Left            =   9240
      Top             =   5400
   End
   Begin VB.Timer TimerTir 
      Interval        =   10
      Left            =   9240
      Top             =   4920
   End
   Begin VB.Timer TimerT 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9240
      Top             =   4440
   End
   Begin VB.Timer Timer 
      Interval        =   10
      Left            =   9240
      Top             =   3960
   End
   Begin VB.Image AST 
      Height          =   690
      Index           =   6
      Left            =   8400
      Picture         =   "Aste.frx":16C5
      Top             =   8280
      Width           =   660
   End
   Begin VB.Image AST 
      Height          =   525
      Index           =   5
      Left            =   7080
      Picture         =   "Aste.frx":1DFF
      Top             =   8160
      Width           =   555
   End
   Begin VB.Image AST 
      Height          =   780
      Index           =   4
      Left            =   5880
      Picture         =   "Aste.frx":23FC
      Top             =   8160
      Width           =   735
   End
   Begin VB.Image AST 
      Height          =   525
      Index           =   3
      Left            =   4680
      Picture         =   "Aste.frx":2D70
      Top             =   8280
      Width           =   630
   End
   Begin VB.Image AST 
      Height          =   1125
      Index           =   2
      Left            =   3120
      Picture         =   "Aste.frx":349F
      Top             =   8040
      Width           =   1050
   End
   Begin VB.Image AST 
      Height          =   765
      Index           =   1
      Left            =   1800
      Picture         =   "Aste.frx":43F4
      Top             =   8280
      Width           =   870
   End
   Begin VB.Image AST 
      Height          =   600
      Index           =   0
      Left            =   360
      Picture         =   "Aste.frx":4C77
      Top             =   8160
      Width           =   720
   End
   Begin VB.Image Image5 
      Height          =   8235
      Left            =   10080
      Picture         =   "Aste.frx":53B8
      Top             =   120
      Width           =   9825
   End
   Begin VB.Image Image4 
      Height          =   8340
      Left            =   10320
      Picture         =   "Aste.frx":D36A
      Top             =   960
      Width           =   10170
   End
   Begin VB.Image Image3 
      Height          =   8160
      Left            =   10320
      Picture         =   "Aste.frx":126A5
      Top             =   2640
      Width           =   10065
   End
   Begin VB.Image Image2 
      Height          =   8280
      Left            =   10320
      Picture         =   "Aste.frx":19384
      Top             =   120
      Width           =   9990
   End
   Begin VB.Image Image1 
      Height          =   8070
      Left            =   10320
      Picture         =   "Aste.frx":1DA6C
      Top             =   -120
      Width           =   9780
   End
   Begin VB.Image AST 
      Height          =   1365
      Index           =   8
      Left            =   7560
      Picture         =   "Aste.frx":23F1B
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Image AST 
      Height          =   465
      Index           =   7
      Left            =   9120
      Picture         =   "Aste.frx":24600
      Top             =   8160
      Width           =   510
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   0
      Left            =   10080
      Picture         =   "Aste.frx":24A45
      Top             =   1920
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   10
      Left            =   9960
      Picture         =   "Aste.frx":24E55
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   9
      Left            =   9960
      Picture         =   "Aste.frx":25265
      Top             =   2280
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   8
      Left            =   10080
      Picture         =   "Aste.frx":25675
      Top             =   3240
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   7
      Left            =   9960
      Picture         =   "Aste.frx":25A85
      Top             =   1200
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   6
      Left            =   10080
      Picture         =   "Aste.frx":25E95
      Top             =   3840
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   5
      Left            =   9960
      Picture         =   "Aste.frx":262A5
      Top             =   720
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   4
      Left            =   10080
      Picture         =   "Aste.frx":266B5
      Top             =   4200
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   3
      Left            =   9840
      Picture         =   "Aste.frx":26AC5
      Top             =   480
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   2
      Left            =   9840
      Picture         =   "Aste.frx":26ED5
      Top             =   2760
      Width           =   180
   End
   Begin VB.Image Feu 
      Height          =   300
      Index           =   1
      Left            =   9960
      Picture         =   "Aste.frx":272E5
      Top             =   4800
      Width           =   180
   End
   Begin VB.Image Ship 
      Height          =   870
      Left            =   4000
      Picture         =   "Aste.frx":276F5
      Top             =   6700
      Width           =   1290
   End
   Begin VB.Label lblCoups 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   7245
      Width           =   495
   End
   Begin VB.Label lblVies 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   6960
      Width           =   495
   End
   Begin VB.Image Stats 
      Height          =   1065
      Left            =   120
      Picture         =   "Aste.frx":27F86
      Top             =   6840
      Width           =   2040
   End
   Begin VB.Image LEVEL 
      Height          =   750
      Index           =   3
      Left            =   1080
      Picture         =   "Aste.frx":292D9
      Top             =   8400
      Width           =   3000
   End
   Begin VB.Image LEVEL 
      Height          =   750
      Index           =   2
      Left            =   6360
      Picture         =   "Aste.frx":2A274
      Top             =   8280
      Width           =   3000
   End
   Begin VB.Image LEVEL 
      Height          =   750
      Index           =   0
      Left            =   5880
      Picture         =   "Aste.frx":2B268
      Top             =   8280
      Width           =   3000
   End
   Begin VB.Image LEVEL 
      Height          =   750
      Index           =   1
      Left            =   3000
      Picture         =   "Aste.frx":2C609
      Top             =   8280
      Width           =   3000
   End
   Begin VB.Image LEVEL 
      Height          =   750
      Index           =   4
      Left            =   3960
      Picture         =   "Aste.frx":2D5EE
      Top             =   8280
      Width           =   3000
   End
End
Attribute VB_Name = "frmAste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim KEY As KeyCodeConstants, KEY2 As KeyCodeConstants
Dim intT As Integer, intFeu As Integer, intVies As Integer
Dim BLNBord As Boolean
Dim intLevel As Integer
Dim IMAGEE As PictureTypeConstants
'****************Merci JCLK****************************************************
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' Joue le fichier de manière synchrone (attend la fin de la lecture pour rendre la main).
Private Const SND_SYNC = &H0
' Joue le fichier de manière asynchrone (rend la main immédiatement).
Private Const SND_ASYNC = &H1
' N'attend pas si le driver son est occupé.
Private Const SND_NOWAIT = &H2000
Private Sub Wav(Fichier As String)
' Joue le fichier son envoyé en paramêtre si le driver est disponible.
Call sndPlaySound(App.Path & Fichier, SND_ASYNC Or SND_NOWAIT)
End Sub
'******************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    If BLNBord = True Then
        Feu(intFeu).Left = Ship.Left + 400
        Feu(intFeu).Top = Ship.Top + 400
        BLNBord = False
    Else
        Feu(intFeu).Left = Ship.Left + 770
        Feu(intFeu).Top = Ship.Top + 400
        BLNBord = True
    End If
    TimerTir.Enabled = True
    intFeu = intFeu + 1
    If intFeu > 10 Then
        intFeu = 0
    End If
    If Len(App.Path) > 3 Then
        Wav ("\Tir.wav")
    Else
        Wav ("Tir.wav")
    End If
Else
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
    KEY = KeyCode
    Timer.Enabled = True
    End If
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
Timer.Enabled = False
End If
intT = TimerT.Enabled = True
End Sub
Private Sub Depart()
Dim intX As Integer
Select Case intLevel
    Case 1
        frmAste.Picture = Image1.Picture
        Call NewLevel(LEVEL(0))
    Case 2
        frmAste.Picture = Image2.Picture
        Call NewLevel(LEVEL(1))
    Case 3
        frmAste.Picture = Image3.Picture
        Call NewLevel(LEVEL(2))
    Case 4
        frmAste.Picture = Image4.Picture
        Call NewLevel(LEVEL(3))
    Case 5
        frmAste.Picture = Image5.Picture
        Call NewLevel(LEVEL(4))
End Select
intFeu = 0
lblVies.Caption = intVies
lblVies.Caption = intVies
Timer.Enabled = False
TimerTir.Enabled = True
TimerTir.Interval = 10
Ship.Left = 4000
Ship.Top = 7000

For intX = 0 To 8
    AST(intX).Visible = True
    AST(intX).Top = 8600
Next
If Len(App.Path) > 3 Then
        Wav ("\Go.wav")
    Else
        Wav ("Go.wav")
    End If
End Sub
Private Sub Form_Load()
intVies = 5
intLevel = 1
Call Depart
End Sub
Private Sub BOUGE()
Select Case KEY
    Case vbKeyUp
        Ship.Top = Ship.Top - 100
    Case vbKeyDown
        Ship.Top = Ship.Top + 100
    Case vbKeyRight
        Ship.Left = Ship.Left + 100
    Case vbKeyLeft
        Ship.Left = Ship.Left - 100
End Select

End Sub

Private Sub imgLevel_Click()

End Sub
Private Sub NewLevel(IMG As IMAGE)

IMG.Top = 3240 + 750
IMG.Left = 3480
TimerLevel.Enabled = True
End Sub
Private Sub Timer_Timer()
Dim blnOk As Boolean

blnOk = True

If (Ship.Left + Ship.Width) < 0 Then
    Ship.Left = (frmAste.Width + Ship.Width)
Else
    If Ship.Left > (frmAste.Width + Ship.Width) Then
    Ship.Left = (0 - Ship.Width)
    End If
End If
If Ship.Top < 0 Then
    Call stopTIMER
    blnOk = False
    Ship.Top = 0
Else
    If Ship.Top > (frmAste.Height - Ship.Height) Then
        Call stopTIMER
    blnOk = False
    Ship.Top = (frmAste.Height - Ship.Height)
    End If
End If
If blnOk = True Then
Call BOUGE
End If


End Sub

Private Sub stopTIMER()
Timer.Enabled = False
End Sub


Private Sub StopTimerT()
TimerT.Enabled = False
End Sub

Private Sub TimerAst_Timer()
Dim intX As Integer, intW As Integer

intX = CInt(Rnd() * 80)

Select Case intX
    Case 0
        If AST(0).Top > 8600 Then
            AST(0).Left = CInt(Rnd() * 9000)
            AST(0).Top = -400
        End If
    Case 1
        If intLevel > 2 Then
            If AST(0).Top > 8600 Then
                AST(0).Left = CInt(Rnd() * 9000)
                AST(0).Top = -400
            End If
        End If
    Case 2
        If intLevel > 4 Then
            If AST(0).Top > 8600 Then
                AST(0).Left = CInt(Rnd() * 9000)
                AST(0).Top = -400
            End If
        End If
    Case 3
        If AST(1).Top > 8600 Then
            AST(1).Left = CInt(Rnd() * 9000)
            AST(1).Top = -400
        End If
    Case 4
        If intLevel > 2 Then
            If AST(1).Top > 8600 Then
                AST(1).Left = CInt(Rnd() * 9000)
                AST(1).Top = -400
            End If
        End If
    Case 5
        If intLevel > 4 Then
            If AST(1).Top > 8600 Then
                AST(1).Left = CInt(Rnd() * 9000)
                AST(1).Top = -400
            End If
        End If
    Case 6
        If AST(2).Top > 8600 Then
            AST(2).Left = CInt(Rnd() * 9000)
            AST(2).Top = -400
        End If
    Case 7
        If intLevel > 2 Then
            If AST(2).Top > 8600 Then
                AST(2).Left = CInt(Rnd() * 9000)
                AST(2).Top = -400
            End If
        End If
    Case 8
        If intLevel > 4 Then
            If AST(2).Top > 8600 Then
                AST(2).Left = CInt(Rnd() * 9000)
                AST(2).Top = -400
            End If
        End If
    Case 9
        If AST(3).Top > 8600 Then
            AST(3).Left = CInt(Rnd() * 9000)
            AST(3).Top = -400
        End If
    Case 10
        If intLevel > 2 Then
            If AST(3).Top > 8600 Then
                AST(3).Left = CInt(Rnd() * 9000)
                AST(3).Top = -400
            End If
        End If
    Case 11
        If intLevel > 4 Then
            If AST(3).Top > 8600 Then
                AST(3).Left = CInt(Rnd() * 9000)
                AST(3).Top = -400
            End If
        End If
    Case 12
        If AST(4).Top > 8600 Then
            AST(4).Left = CInt(Rnd() * 9000)
            AST(4).Top = -400
        End If
    Case 13
        If intLevel > 2 Then
            If AST(4).Top > 8600 Then
                AST(4).Left = CInt(Rnd() * 9000)
                AST(4).Top = -400
            End If
        End If
    Case 14
        If intLevel > 4 Then
            If AST(4).Top > 8600 Then
                AST(4).Left = CInt(Rnd() * 9000)
                AST(4).Top = -400
            End If
        End If
    Case 15
        If AST(5).Top > 8600 Then
            AST(5).Left = CInt(Rnd() * 9000)
            AST(5).Top = -400
        End If
    Case 16
        If intLevel > 2 Then
            If AST(5).Top > 8600 Then
                AST(5).Left = CInt(Rnd() * 9000)
                AST(5).Top = -400
            End If
        End If
    Case 17
        If intLevel > 4 Then
            If AST(5).Top > 8600 Then
                AST(5).Left = CInt(Rnd() * 9000)
                AST(5).Top = -400
            End If
        End If
    Case 18
        If AST(6).Top > 8600 Then
            AST(6).Left = CInt(Rnd() * 9000)
            AST(6).Top = -400
        End If
    Case 19
        If intLevel > 2 Then
            If AST(6).Top > 8600 Then
                AST(6).Left = CInt(Rnd() * 9000)
                AST(6).Top = -400
            End If
        End If
    Case 20
        If intLevel > 4 Then
            If AST(6).Top > 8600 Then
                AST(6).Left = CInt(Rnd() * 9000)
                AST(6).Top = -400
            End If
        End If
    Case 21
        If AST(7).Top > 8600 Then
            AST(7).Left = CInt(Rnd() * 9000)
            AST(7).Top = -400
        End If
    Case 22
        If AST(8).Top > 8600 Then
            AST(8).Left = CInt(Rnd() * 9000)
            AST(8).Top = -400
        End If
End Select

End Sub

Private Sub TimerLevel_Timer()
If LEVEL(intLevel - 1).Top > 3240 Then
    LEVEL(intLevel - 1).Top = LEVEL(intLevel - 1).Top - 10
Else
    LEVEL(intLevel - 1).Height = LEVEL(intLevel - 1).Height - 10
    If LEVEL(intLevel - 1).Height < 10 Then
        LEVEL(intLevel - 1).Top = 10000
        LEVEL(intLevel - 1).Height = 750
        Call StopTimerLevel
    End If
End If
End Sub
Private Sub StopTimerLevel()
TimerLevel.Enabled = False
End Sub

Private Sub TimerT_Timer()
intT = intT - 1
Select Case KEY
    Case vbKeyUp
        Ship.Top = Ship.Top - intT
    Case vbKeyDown
        Ship.Top = Ship.Top + intT
    Case vbKeyRight
        Ship.Left = Ship.Left + intT
    Case vbKeyLeft
        Ship.Left = Ship.Left - intT
End Select
If intT = 0 Then
Call StopTimerT
End If
End Sub

Private Sub TimerTir_Timer()
Dim intX As Integer, intW As Integer
For intX = 0 To 10
    Feu(intX).Top = Feu(intX).Top - 100
    For intW = 0 To 6
        If Feu(intX).Top > AST(intW).Top And Feu(intX).Top < (AST(intW).Top + AST(intW).Height) And Feu(intX).Left > AST(intW).Left And Feu(intX).Left < (AST(intW).Left + AST(intW).Width) Then
            lblCoups.Caption = CStr(Val(lblCoups.Caption + 1))
            AST(intW).Top = 8600
            If Len(App.Path) > 3 Then
                Wav ("\Pogne.wav")
            Else
                Wav ("Pogne.wav")
            End If
            Select Case intLevel
                Case 1
                    If Val(lblCoups.Caption) = 60 Then
                        intLevel = 2
                        Call Depart
                    End If
                Case 2
                    If Val(lblCoups.Caption) = 120 Then
                        intLevel = 3
                        Call Depart
                    End If
                Case 3
                    If Val(lblCoups.Caption) = 180 Then
                        intLevel = 4
                        Call Depart
                    End If
                Case 4
                    If Val(lblCoups.Caption) = 240 Then
                        intLevel = 5
                        Call Depart
                    End If
                Case 5
                    If Val(lblCoups.Caption) = 300 Then
                        intLevel = 1
                        lblCoups.Caption = 0
                        intVies = 5
                        If MsgBox("Bravo! Vous êtes un grand pilote, vous avez fait " & lblCoups.Caption & " points. Rejouer?", vbYesNo, "Astéroïde!") = vbYes Then
                            Call Depart
                        Else
                        Unload frmAste
                        End If
                    End If
            End Select
        End If
    Next
Next

End Sub
Private Sub BOUGEAST(V As Integer)
Dim intX As Integer, intW As Integer
For intX = 0 To 8
If intX <> 8 Then
    AST(intX).Top = AST(intX).Top + CInt(intLevel * V / 3)
    If intX <> 7 Then
         If AST(intX).Top > Ship.Top And AST(intX).Top < (Ship.Top + Ship.Height) And AST(intX).Left > Ship.Left And AST(intX).Left < (Ship.Left + Ship.Width) Then
            AST(intX).Top = 8600
            If Len(App.Path) > 3 Then
                Wav ("\Crash.wav")
            Else
                Wav ("Crash.wav")
            End If
            If Val(lblVies.Caption) > 0 Then
                lblVies.Caption = CStr(Val(lblVies.Caption - 1))
            Else
                If MsgBox("Partie terminée! Vous avez détruit " & lblCoups.Caption & " astéroid(s). Rejouer?", vbYesNo, "Astéroïd!") = vbYes Then
                    intLevel = 1
                    lblCoups.Caption = 0
                    Call Depart
                Else
                    Unload frmAste
                End If
            End If
        End If
    Else
        If AST(intX).Top > Ship.Top And AST(intX).Top < (Ship.Top + Ship.Height) And AST(intX).Left > Ship.Left And AST(intX).Left < (Ship.Left + Ship.Width) Then
            AST(intX).Top = 8600
            If Len(App.Path) > 3 Then
                Wav ("\Vie.wav")
            Else
                Wav ("Vie.wav")
            End If
            lblVies.Caption = CStr(Val(lblVies.Caption + 1))
        End If
    End If
    If intX <> 7 Then
         If AST(intX).Top > frmAste.Height And AST(intX).Top < (frmAste.Height + 200) Then
            AST(intX).Top = 8600
            If Val(lblCoups.Caption) > 0 Then
                lblCoups.Caption = CStr(Val(lblCoups.Caption - 1))
            End If
        End If
    End If
Else
    AST(intX).Top = AST(intX).Top + (3 * V)
    AST(intX).Left = AST(intX).Left - (1.5 * V)
    If AST(intX).Top > Ship.Top And AST(intX).Top < (Ship.Top + Ship.Height) And AST(intX).Left > Ship.Left And AST(intX).Left < (Ship.Left + Ship.Width) Then
            AST(intX).Top = 8600
            If Len(App.Path) > 3 Then
                Wav ("\Crash.wav")
            Else
                Wav ("Crash.wav")
            End If
            If Val(lblVies.Caption) > 0 Then
                lblVies.Caption = CStr(Val(lblVies.Caption - 1))
            Else
                If MsgBox("Partie terminée! Vous avez détruit " & lblCoups.Caption & " astéroid(s). Rejouer?", vbYesNo, "Astéroïd!") = vbYes Then
                    Call Depart
                Else
                    Unload frmAste
                End If
            End If
        End If
End If
Next
       
End Sub

Private Sub TV0_Timer()
Dim intX As Integer

    Call BOUGEAST(100)
    
End Sub
