VERSION 5.00
Begin VB.Form niveau2 
   BackColor       =   &H8000000D&
   Caption         =   "Arkanoïd"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   8295
      Left            =   0
      Picture         =   "niveau2.frx":0000
      ScaleHeight     =   8235
      ScaleWidth      =   7755
      TabIndex        =   28
      Top             =   0
      Width           =   7815
      Begin VB.Image Image1 
         Height          =   270
         Left            =   3360
         Picture         =   "niveau2.frx":36CBE
         Top             =   7800
         Width           =   1125
      End
      Begin VB.Image img2Niv1 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":370A5
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image img3Niv1 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":37436
         Top             =   1410
         Width           =   645
      End
      Begin VB.Image img4Niv1 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":377C7
         Top             =   1410
         Width           =   645
      End
      Begin VB.Image img5Niv1 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":37B58
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image img6Niv1 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":37EE9
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image img7Niv1 
         Height          =   270
         Left            =   2580
         Picture         =   "niveau2.frx":3827A
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image img8Niv1 
         Height          =   270
         Left            =   3870
         Picture         =   "niveau2.frx":3860B
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image img9Niv1 
         Height          =   270
         Left            =   2580
         Picture         =   "niveau2.frx":3899C
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image img10Niv1 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":38D2D
         Top             =   1140
         Width           =   645
      End
      Begin VB.Image img11Niv1 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":390BE
         Top             =   1140
         Width           =   645
      End
      Begin VB.Image img12Niv1 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":3944F
         Top             =   1410
         Width           =   645
      End
      Begin VB.Image img13Niv1 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":397E0
         Top             =   1140
         Width           =   645
      End
      Begin VB.Image imgBall 
         Height          =   240
         Left            =   3850
         Picture         =   "niveau2.frx":39B71
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   240
      End
      Begin VB.Image img25Niv1 
         Height          =   270
         Left            =   3240
         Picture         =   "niveau2.frx":39E7B
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image img24Niv1 
         Height          =   270
         Left            =   4515
         Picture         =   "niveau2.frx":3A20C
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image img23Niv1 
         Height          =   270
         Left            =   5850
         Picture         =   "niveau2.frx":3A59D
         Top             =   4080
         Width           =   645
      End
      Begin VB.Image img22Niv1 
         Height          =   270
         Left            =   3870
         Picture         =   "niveau2.frx":3A92E
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image img21Niv1 
         Height          =   270
         Left            =   5205
         Picture         =   "niveau2.frx":3ACBF
         Top             =   4080
         Width           =   645
      End
      Begin VB.Image img20Niv1 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":3B050
         Top             =   330
         Width           =   645
      End
      Begin VB.Image img19Niv1 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":3B3E1
         Top             =   1140
         Width           =   645
      End
      Begin VB.Image img18Niv1 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":3B772
         Top             =   600
         Width           =   645
      End
      Begin VB.Image img17Niv1 
         Height          =   270
         Left            =   4560
         Picture         =   "niveau2.frx":3BB03
         Top             =   4080
         Width           =   645
      End
      Begin VB.Image img16Niv1 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":3BE94
         Top             =   870
         Width           =   645
      End
      Begin VB.Image img15Niv1 
         Height          =   270
         Left            =   2580
         Picture         =   "niveau2.frx":3C225
         Top             =   1410
         Width           =   645
      End
      Begin VB.Image img14Niv1 
         Height          =   270
         Left            =   5160
         Picture         =   "niveau2.frx":3C5B6
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image5 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":3C947
         Top             =   870
         Width           =   645
      End
      Begin VB.Image Image6 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":3CCD8
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image Image7 
         Height          =   270
         Left            =   3240
         Picture         =   "niveau2.frx":3D069
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image Image8 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":3D3FA
         Top             =   1410
         Width           =   645
      End
      Begin VB.Image Image9 
         Height          =   270
         Left            =   4515
         Picture         =   "niveau2.frx":3D78B
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image10 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":3DB1C
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image11 
         Height          =   270
         Left            =   3225
         Picture         =   "niveau2.frx":3DEAD
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image Image12 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":3E23E
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image13 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":3E5CF
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image Image14 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":3E960
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image Image15 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":3ECF1
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image Image16 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":3F082
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image Image17 
         Height          =   270
         Left            =   2580
         Picture         =   "niveau2.frx":3F413
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image Image18 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":3F7A4
         Top             =   600
         Width           =   645
      End
      Begin VB.Image Image19 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":3FB35
         Top             =   870
         Width           =   645
      End
      Begin VB.Image Image20 
         Height          =   270
         Left            =   3870
         Picture         =   "niveau2.frx":3FEC6
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image21 
         Height          =   270
         Left            =   3225
         Picture         =   "niveau2.frx":40257
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image22 
         Height          =   270
         Left            =   2580
         Picture         =   "niveau2.frx":405E8
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image23 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":40979
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image24 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":40D0A
         Top             =   2490
         Width           =   645
      End
      Begin VB.Image Image25 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":4109B
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image Image26 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":4142C
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image Image27 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":417BD
         Top             =   1950
         Width           =   645
      End
      Begin VB.Image Image28 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":41B4E
         Top             =   2220
         Width           =   645
      End
      Begin VB.Image Image29 
         Height          =   270
         Left            =   0
         Picture         =   "niveau2.frx":41EDF
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image30 
         Height          =   270
         Left            =   645
         Picture         =   "niveau2.frx":42270
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image31 
         Height          =   270
         Left            =   1290
         Picture         =   "niveau2.frx":42601
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image32 
         Height          =   270
         Left            =   1935
         Picture         =   "niveau2.frx":42992
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image33 
         Height          =   270
         Left            =   2580
         Picture         =   "niveau2.frx":42D23
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image34 
         Height          =   270
         Left            =   3225
         Picture         =   "niveau2.frx":430B4
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image35 
         Height          =   270
         Left            =   3885
         Picture         =   "niveau2.frx":43445
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image36 
         Height          =   270
         Left            =   4515
         Picture         =   "niveau2.frx":437D6
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image37 
         Height          =   270
         Left            =   5160
         Picture         =   "niveau2.frx":43B67
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image38 
         Height          =   270
         Left            =   5805
         Picture         =   "niveau2.frx":43EF8
         Top             =   2760
         Width           =   645
      End
      Begin VB.Image Image39 
         Height          =   270
         Left            =   6495
         Picture         =   "niveau2.frx":44289
         Top             =   4080
         Width           =   645
      End
      Begin VB.Image Image40 
         Height          =   270
         Left            =   7140
         Picture         =   "niveau2.frx":4461A
         Top             =   4080
         Width           =   645
      End
   End
   Begin VB.Timer timerDepBall 
      Left            =   7920
      Top             =   7800
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
      Left            =   8040
      TabIndex        =   25
      Text            =   "0"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox txtBrique 
      Height          =   285
      Left            =   7920
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8640
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7920
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9000
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7920
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8280
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8280
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8640
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   9000
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   8280
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   9000
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   7920
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   8280
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   7920
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8640
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   8280
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   9000
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   8640
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   9000
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   8280
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   9000
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   8640
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtVie 
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Text            =   "3"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   8040
      Picture         =   "niveau2.frx":449AB
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   645
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   8760
      Picture         =   "niveau2.frx":44D92
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   9480
      Picture         =   "niveau2.frx":45179
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   645
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
      Left            =   8040
      TabIndex        =   27
      Top             =   5520
      Width           =   1455
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
      Left            =   8040
      TabIndex        =   26
      Top             =   6360
      Width           =   1455
   End
End
Attribute VB_Name = "niveau2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nblancement As Integer
Public posDepL As Integer
Public posDepT As Integer
Public direction As Integer
Public sensVert As String

Private Sub Form_Load()
    posDepL = 25
    posDepT = 25
    direction = 2
    sensVert = "haut"
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
                timerDepBall.Interval = 1
                nblancement = 1
            End If
    End Select
    If txtSCORE.Text = 8400 Then
        niveau2.txtSCORE.Text = 8400
        niveau2.txtVie.Text = txtVie.Text
        Unload Me
        niveau2.Show vbModal
    End If
End Sub

Private Sub timerDepBall_Timer()
    Call deplacmentBall
    Call orientationBall
    Call verifEliminationBrique
    'Call verifEliminationBrique2
    'Call verifEliminationBrique3
End Sub

Sub deplacmentBall()
'---Gestion des déplacemnt de la ball---
    If direction = 1 Then
        imgBall.Left = imgBall.Left + posDepL
        imgBall.Top = imgBall.Top + posDepT
        sensVert = "bas"
    ElseIf direction = 2 Then
        imgBall.Left = imgBall.Left - posDepL
        imgBall.Top = imgBall.Top - posDepT
        sensVert = "haut"
    ElseIf direction = 3 Then
        imgBall.Left = imgBall.Left - posDepL
        imgBall.Top = imgBall.Top + posDepT
        sensVert = "bas"
    ElseIf direction = 4 Then
        imgBall.Left = imgBall.Left + posDepL
        imgBall.Top = imgBall.Top - posDepT
        sensVert = "haut"
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
Sub verifEliminationBrique2()

    If imgBall.Left >= img14Niv1.Left And imgBall.Left <= img14Niv1.Left + img14Niv1.Width Then
            If imgBall.Top <= img14Niv1.Top + img14Niv1.Height And img14Niv1.Visible = True Then
                Call newdirection
                If Text1.Text = "brique" Then
                    img14Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text1.Text = "brique"
            End If
    End If
    If imgBall.Left >= img15Niv1.Left And imgBall.Left <= img15Niv1.Left + img15Niv1.Width Then
            If imgBall.Top <= img15Niv1.Top + img15Niv1.Height And img15Niv1.Visible = True Then
                Call newdirection
                If Text2.Text = "brique" Then
                    img15Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text2.Text = "brique"
            End If
    End If
    If imgBall.Left >= img16Niv1.Left And imgBall.Left <= img16Niv1.Left + img16Niv1.Width Then
            If imgBall.Top <= img16Niv1.Top + img16Niv1.Height And img16Niv1.Visible = True Then
                Call newdirection
                If Text3.Text = "brique" Then
                    img16Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text3.Text = "brique"
            End If
    End If
    If imgBall.Left >= img17Niv1.Left And imgBall.Left <= img17Niv1.Left + img17Niv1.Width Then
            If imgBall.Top <= img17Niv1.Top + img17Niv1.Height And img17Niv1.Visible = True Then
                Call newdirection
                If Text4.Text = "brique" Then
                    img17Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text4.Text = "brique"
            End If
    End If
    If imgBall.Left >= img18Niv1.Left And imgBall.Left <= img18Niv1.Left + img18Niv1.Width Then
            If imgBall.Top <= img18Niv1.Top + img18Niv1.Height And img18Niv1.Visible = True Then
                Call newdirection
                If txtBrique.Text = "brique" Then
                    img18Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                txtBrique.Text = "brique"
            End If
    End If
    If imgBall.Left >= img19Niv1.Left And imgBall.Left <= img19Niv1.Left + img19Niv1.Width Then
            If imgBall.Top <= img19Niv1.Top + img19Niv1.Height And img19Niv1.Visible = True Then
                Call newdirection
                If Text5.Text = "brique" Then
                    img19Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text5.Text = "brique"
            End If
    End If
    If imgBall.Left >= img20Niv1.Left And imgBall.Left <= img20Niv1.Left + img20Niv1.Width Then
            If imgBall.Top <= img20Niv1.Top + img20Niv1.Height And img20Niv1.Visible = True Then
                Call newdirection
                If Text6.Text = "brique" Then
                    img20Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text6.Text = "brique"
            End If
    End If
    If imgBall.Left >= img21Niv1.Left And imgBall.Left <= img21Niv1.Left + img21Niv1.Width Then
            If imgBall.Top <= img21Niv1.Top + img21Niv1.Height And img21Niv1.Visible = True Then
                Call newdirection
                If Text7.Text = "brique" Then
                    img21Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text7.Text = "brique"
            End If
    End If
    If imgBall.Left >= img22Niv1.Left And imgBall.Left <= img22Niv1.Left + img22Niv1.Width Then
            If imgBall.Top <= img22Niv1.Top + img22Niv1.Height And img22Niv1.Visible = True Then
                Call newdirection
                If Text8.Text = "brique" Then
                    img22Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text8.Text = "brique"
            End If
    End If
    If imgBall.Left >= img23Niv1.Left And imgBall.Left <= img23Niv1.Left + img23Niv1.Width Then
            If imgBall.Top <= img23Niv1.Top + img23Niv1.Height And img23Niv1.Visible = True Then
                Call newdirection
                If Text9.Text = "brique" Then
                    img23Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text9.Text = "brique"
            End If
    End If
    If imgBall.Left >= img24Niv1.Left And imgBall.Left <= img24Niv1.Left + img24Niv1.Width Then
            If imgBall.Top <= img24Niv1.Top + img24Niv1.Height And img24Niv1.Visible = True Then
                Call newdirection
                If Text10.Text = "brique" Then
                    img24Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text10.Text = "brique"
            End If
    End If
    If imgBall.Left >= img25Niv1.Left And imgBall.Left <= img25Niv1.Left + img25Niv1.Width Then
            If imgBall.Top <= img25Niv1.Top + img25Niv1.Height And img25Niv1.Visible = True Then
                Call newdirection
                If Text11.Text = "brique" Then
                    img25Niv1.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text11.Text = "brique"
            End If
    End If
End Sub

Sub verifEliminationBrique()
    Call eliminationBrique(img2Niv1)
    Call eliminationBrique(img3Niv1)
    Call eliminationBrique(img4Niv1)
    Call eliminationBrique(img5Niv1)
    Call eliminationBrique(img6Niv1)
    Call eliminationBrique(img7Niv1)
    Call eliminationBrique(img8Niv1)
    Call eliminationBrique(img9Niv1)
    Call eliminationBrique(img10Niv1)
    Call eliminationBrique(img11Niv1)
    Call eliminationBrique(img12Niv1)
    Call eliminationBrique(img13Niv1)
    Call eliminationBrique(Image5)
    Call eliminationBrique(Image6)
    Call eliminationBrique(Image7)
    Call eliminationBrique(Image8)
    Call eliminationBrique(Image9)
    Call eliminationBrique(Image10)
    Call eliminationBrique(Image11)
    Call eliminationBrique(Image12)
    Call eliminationBrique(Image13)
    Call eliminationBrique(Image14)
    Call eliminationBrique(Image15)
    Call eliminationBrique(Image16)
    Call eliminationBrique(Image17)
    Call eliminationBrique(Image18)
    Call eliminationBrique(Image19)
    Call eliminationBrique(Image20)
    Call eliminationBrique(Image21)
    Call eliminationBrique(Image22)
    Call eliminationBrique(Image23)
    Call eliminationBrique(Image24)
    Call eliminationBrique(Image25)
    Call eliminationBrique(Image26)
    Call eliminationBrique(Image27)
    Call eliminationBrique(Image28)
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

Sub verifEliminationBrique3()

    If imgBall.Left >= Image29.Left And imgBall.Left <= Image29.Left + Image29.Width Then
            If imgBall.Top <= Image29.Top + Image29.Height And Image29.Visible = True Then
                Call newdirection
                If Text1.Text = "brique" Then
                    Image29.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text1.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image30.Left And imgBall.Left <= Image30.Left + Image30.Width Then
            If imgBall.Top <= Image30.Top + Image30.Height And Image30.Visible = True Then
                Call newdirection
                If Text2.Text = "brique" Then
                    Image30.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text2.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image31.Left And imgBall.Left <= Image31.Left + Image31.Width Then
            If imgBall.Top <= Image31.Top + Image31.Height And Image31.Visible = True Then
                Call newdirection
                If Text3.Text = "brique" Then
                    Image31.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text3.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image32.Left And imgBall.Left <= Image32.Left + Image32.Width Then
            If imgBall.Top <= Image32.Top + Image32.Height And Image32.Visible = True Then
                Call newdirection
                If Text4.Text = "brique" Then
                    Image32.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text4.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image33.Left And imgBall.Left <= Image33.Left + Image33.Width Then
            If imgBall.Top <= Image33.Top + Image33.Height And Image33.Visible = True Then
                Call newdirection
                If txtBrique.Text = "brique" Then
                    Image33.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                txtBrique.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image34.Left And imgBall.Left <= Image34.Left + Image34.Width Then
            If imgBall.Top <= Image34.Top + Image34.Height And Image34.Visible = True Then
                Call newdirection
                If Text5.Text = "brique" Then
                    Image34.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text5.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image35.Left And imgBall.Left <= Image35.Left + Image35.Width Then
            If imgBall.Top <= Image35.Top + Image35.Height And Image35.Visible = True Then
                Call newdirection
                If Text6.Text = "brique" Then
                    Image35.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text6.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image36.Left And imgBall.Left <= Image36.Left + Image36.Width Then
            If imgBall.Top <= Image36.Top + Image36.Height And Image36.Visible = True Then
                Call newdirection
                If Text7.Text = "brique" Then
                    Image36.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text7.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image37.Left And imgBall.Left <= Image37.Left + Image37.Width Then
            If imgBall.Top <= Image37.Top + Image37.Height And Image37.Visible = True Then
                Call newdirection
                If Text8.Text = "brique" Then
                    Image37.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text8.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image38.Left And imgBall.Left <= Image38.Left + Image38.Width Then
            If imgBall.Top <= Image38.Top + Image38.Height And Image38.Visible = True Then
                Call newdirection
                If Text9.Text = "brique" Then
                    Image38.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text9.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image39.Left And imgBall.Left <= Image39.Left + Image39.Width Then
            If imgBall.Top <= Image39.Top + Image39.Height And Image39.Visible = True Then
                Call newdirection
                If Text10.Text = "brique" Then
                    Image39.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text10.Text = "brique"
            End If
    End If
    If imgBall.Left >= Image40.Left And imgBall.Left <= Image40.Left + Image40.Width Then
            If imgBall.Top <= Image40.Top + Image40.Height And Image40.Visible = True Then
                Call newdirection
                If Text11.Text = "brique" Then
                    Image40.Visible = False
                    txtSCORE.Text = txtSCORE.Text + 200
                End If
                Text11.Text = "brique"
            End If
    End If
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

