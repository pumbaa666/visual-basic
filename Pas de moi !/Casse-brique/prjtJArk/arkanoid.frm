VERSION 5.00
Begin VB.Form arkanoid 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARKANOÏD"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVie 
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Text            =   "3"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   8760
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   9120
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   8400
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   8040
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   9120
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   9120
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   8400
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8760
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   8040
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   8400
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   8040
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   9120
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   8400
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   9120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8760
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8400
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8400
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8760
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8040
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8040
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBrique 
      Height          =   285
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSCORE 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8160
      TabIndex        =   1
      Text            =   "0"
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Timer timerDepBall 
      Left            =   7920
      Top             =   7800
   End
   Begin VB.PictureBox Picture1 
      Height          =   8295
      Left            =   0
      Picture         =   "arkanoid.frx":0000
      ScaleHeight     =   8235
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   35
         Left            =   7080
         Picture         =   "arkanoid.frx":3CCA2
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   34
         Left            =   6450
         Picture         =   "arkanoid.frx":3D033
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   33
         Left            =   5805
         Picture         =   "arkanoid.frx":3D3C4
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   32
         Left            =   5160
         Picture         =   "arkanoid.frx":3D755
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   31
         Left            =   4515
         Picture         =   "arkanoid.frx":3DAE6
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   30
         Left            =   3870
         Picture         =   "arkanoid.frx":3DE77
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   29
         Left            =   3225
         Picture         =   "arkanoid.frx":3E208
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   28
         Left            =   2580
         Picture         =   "arkanoid.frx":3E599
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   27
         Left            =   1935
         Picture         =   "arkanoid.frx":3E92A
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   26
         Left            =   1290
         Picture         =   "arkanoid.frx":3ECBB
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   25
         Left            =   0
         Picture         =   "arkanoid.frx":3F04C
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   24
         Left            =   645
         Picture         =   "arkanoid.frx":3F3DD
         Top             =   1710
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   23
         Left            =   1290
         Picture         =   "arkanoid.frx":3F76E
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   22
         Left            =   0
         Picture         =   "arkanoid.frx":3FAFF
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   21
         Left            =   645
         Picture         =   "arkanoid.frx":3FE90
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   20
         Left            =   1935
         Picture         =   "arkanoid.frx":40221
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   19
         Left            =   2580
         Picture         =   "arkanoid.frx":405B2
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   18
         Left            =   3225
         Picture         =   "arkanoid.frx":40943
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   17
         Left            =   3870
         Picture         =   "arkanoid.frx":40CD4
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   16
         Left            =   4515
         Picture         =   "arkanoid.frx":41065
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   15
         Left            =   5160
         Picture         =   "arkanoid.frx":413F6
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   14
         Left            =   5805
         Picture         =   "arkanoid.frx":41787
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   13
         Left            =   6450
         Picture         =   "arkanoid.frx":41B18
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   12
         Left            =   7095
         Picture         =   "arkanoid.frx":41EA9
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   11
         Left            =   7095
         Picture         =   "arkanoid.frx":4223A
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   10
         Left            =   6450
         Picture         =   "arkanoid.frx":425CB
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   9
         Left            =   5805
         Picture         =   "arkanoid.frx":4295C
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   8
         Left            =   5160
         Picture         =   "arkanoid.frx":42CED
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   7
         Left            =   4515
         Picture         =   "arkanoid.frx":4307E
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   6
         Left            =   3870
         Picture         =   "arkanoid.frx":4340F
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   5
         Left            =   3225
         Picture         =   "arkanoid.frx":437A0
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   4
         Left            =   2580
         Picture         =   "arkanoid.frx":43B31
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   3
         Left            =   1935
         Picture         =   "arkanoid.frx":43EC2
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   2
         Left            =   645
         Picture         =   "arkanoid.frx":44253
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   1
         Left            =   0
         Picture         =   "arkanoid.frx":445E4
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgBrique 
         Height          =   270
         Index           =   0
         Left            =   1290
         Picture         =   "arkanoid.frx":44975
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " LEVEL 1 CLEAR  CLICK HERE TO START  THE NEXT LEVEL"
         BeginProperty Font 
            Name            =   "PanRoman"
            Size            =   18
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   2520
         TabIndex        =   29
         Top             =   3720
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Image imgBall 
         Height          =   240
         Left            =   3840
         Picture         =   "arkanoid.frx":44D06
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   3360
         Picture         =   "arkanoid.frx":45010
         Top             =   7800
         Width           =   1125
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LIFE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   28
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   27
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   9600
      Picture         =   "arkanoid.frx":453F7
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   645
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   8880
      Picture         =   "arkanoid.frx":457DE
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   8160
      Picture         =   "arkanoid.frx":45BC5
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   645
   End
End
Attribute VB_Name = "arkanoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nblancement As Integer
Public posDepL As Integer
Public posDepT As Integer
Public direction As Integer

Private Sub Form_Load()
    posDepL = 25
    posDepT = 25
    direction = 2
End Sub

Private Sub Label3_Click()
    Unload Me
    niveau2.Show vbModal
End Sub

Private Sub picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 37
            Image1.Left = Image1.Left - 400
            If nblancement <> 1 Then
                imgBall.Left = Image1.Left + 380
            End If
        Case Is = 39
            Image1.Left = Image1.Left + 400
            If nblancement <> 1 Then
                imgBall.Left = Image1.Left + 380
            End If
        Case Is = 32
            If nblancement <> 1 Then
                Select Case imgBall.Left
                    Case Is < (Picture1.Width / 2)
                        direction = 2
                    Case Is > (Picture1.Width / 2)
                        direction = 4
                End Select
                timerDepBall.Interval = 1
                nblancement = 1
            End If
    End Select
End Sub

Private Sub timerDepBall_Timer()
    Call deplacmentBall
    Call orientationBall
    Call verifEliminationBrique
    If txtSCORE.Text = 3600 Then
        timerDepBall.Interval = 0
        niveau2.txtSCORE.Text = 3600
        niveau2.txtVie.Text = txtVie.Text
        Label3.Visible = True
    End If
End Sub

Sub deplacmentBall()
'---Gestion des déplacemnt de la ball---
    If direction = 1 Then
        imgBall.Left = imgBall.Left + posDepL
        imgBall.Top = imgBall.Top + posDepT
    ElseIf direction = 2 Then
        imgBall.Left = imgBall.Left - posDepL
        imgBall.Top = imgBall.Top - posDepT
    ElseIf direction = 3 Then
        imgBall.Left = imgBall.Left - posDepL
        imgBall.Top = imgBall.Top + posDepT
    ElseIf direction = 4 Then
        imgBall.Left = imgBall.Left + posDepL
        imgBall.Top = imgBall.Top - posDepT
    End If
End Sub

Sub orientationBall()
    '---Gestion de l'orientation---
    If imgBall.Left < 2 Then
        Select Case direction
            Case Is = 3
                direction = 1
            Case Is = 2
                direction = 4
        End Select
    End If
    If imgBall.Top > Picture1.Height - imgBall.Height - 400 Then
        If imgBall.Left >= Image1.Left - 150 And imgBall.Left <= Image1.Left + Image1.Width Then
                posDepL = Int(Rnd * 50)
                Select Case direction
                    Case Is = 1
                        direction = 4
                    Case Is = 3
                        direction = 2
                End Select
            Else
                timerDepBall.Interval = 0
                imgBall.Visible = False
                txtVie.Text = txtVie.Text - 1
                Call verifVie
        End If
    End If
    If imgBall.Left + imgBall.Width > Picture1.Width - 20 Then
        Select Case direction
            Case Is = 1
                direction = 3
            Case Is = 4
                direction = 2
        End Select
    End If
    If imgBall.Top < 2 Then
        Select Case direction
            Case Is = 2
                direction = 3
            Case Is = 4
                direction = 1
        End Select
    End If
End Sub

'    If imgBall.Left >= img25Niv1.Left And imgBall.Left <= img25Niv1.Left + img25Niv1.Width Then
'            If imgBall.Top <= img25Niv1.Top + img25Niv1.Height And img25Niv1.Visible = True Then
'                Call newdirection
'                If Text11.Text = "brique" Then
'                    img25Niv1.Visible = False
'                    txtSCORE.Text = txtSCORE.Text + 200
'                End If
'                Text11.Text = "brique"
'            End If
'    End If

Sub verifEliminationBrique()
Dim i As Integer
    For i = 0 To 35
        Call eliminationBrique(imgBrique(i))
    Next
End Sub
    
Sub eliminationBrique(image As image)

'---Verifcation si la balle touche la brique par le bas ---
If direction = 2 Or 4 Then
    If image.Visible = True Then
        If imgBall.Top < (image.Top + image.Height) And imgBall.Top >= image.Top - (image.Height / 2) Then
            If imgBall.Left >= image.Left And imgBall.Left <= (image.Left + image.Width) Then
                image.Visible = False
                txtSCORE.Text = txtSCORE.Text + 100
                Call newdirectionhaut
            End If
        End If
    End If
End If
'---Verification si la balle touche la brique par le haut---
If direction = 1 Or 3 Then
    If image.Visible = True Then
        If (imgBall.Top + imgBall.Height) > image.Top And (imgBall.Top + imgBall.Height) <= image.Top + (image.Height / 2) Then
            If imgBall.Left >= image.Left And imgBall.Left <= (image.Left + image.Width) Then
                image.Visible = False
                txtSCORE.Text = txtSCORE.Text + 100
                Call newdirectionbas
            End If
        End If
    End If
End If
 
End Sub
Sub newdirectionbas()
    Select Case direction
        Case Is = 1
            direction = 4
        Case Is = 3
            direction = 2
    End Select
End Sub
Sub newdirectionhaut()
    Select Case direction
        Case Is = 2
            direction = 3
        Case Is = 4
            direction = 1
    End Select
End Sub

Sub verifVie()
Select Case txtVie.Text
    Case Is = "2"
        Image4.Visible = False
        Image1.Left = 3360
        imgBall.Left = 3850
        imgBall.Top = 7560
        imgBall.Visible = True
        nblancement = 0
        Call Form_Load
    Case Is = "1"
        Image3.Visible = False
        Image1.Left = 3360
        imgBall.Left = 3850
        imgBall.Top = 7560
        imgBall.Visible = True
        nblancement = 0
        Call Form_Load
    Case Is = "0"
        Image2.Visible = False
        Image1.Left = 3360
        imgBall.Left = 3850
        imgBall.Top = 7560
        imgBall.Visible = True
        nblancement = 0
        Call Form_Load
    Case Is = "-1"
        MsgBox "Game Over"
End Select
End Sub


Sub newdirection()
If direction = 1 Then
    direction = 4
Else
    direction = 2
End If
If direction = 1 Then
    direction = 3
Else
    direction = 2
End If
If direction = 2 Then
    direction = 3
Else
    direction = 1
End If
If direction = 2 Then
    direction = 4
Else
    direction = 1
End If
End Sub
