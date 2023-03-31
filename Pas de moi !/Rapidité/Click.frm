VERSION 5.00
Begin VB.Form Click 
   BackColor       =   &H00A57A5A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Click"
   ClientHeight    =   5796
   ClientLeft      =   42
   ClientTop       =   336
   ClientWidth     =   10836
   ClipControls    =   0   'False
   Icon            =   "Click.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   10836
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Left            =   9960
      Top             =   3720
   End
   Begin VB.CommandButton CommenceJ3 
      BackColor       =   &H00A78176&
      Caption         =   "Jouer / continuer"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Tape 
      BackColor       =   &H80000001&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   9120
      TabIndex        =   54
      Top             =   4200
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Left            =   6480
      Top             =   5040
   End
   Begin VB.Timer Timer3 
      Left            =   5880
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Left            =   6840
      Top             =   4440
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00A78176&
      Caption         =   "Quitter"
      Height          =   255
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CommenceJ2 
      BackColor       =   &H00A78176&
      Caption         =   "Jouer / continuer"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton B2 
      BackColor       =   &H00A78176&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton B3 
      BackColor       =   &H00A78176&
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton B4 
      BackColor       =   &H00A78176&
      Enabled         =   0   'False
      Height          =   255
      Left            =   5400
      MaskColor       =   &H00A57A5A&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton B1 
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Clickeur 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      Picture         =   "Click.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A78176&
      Caption         =   "Jouer / continuer"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   4200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00A57A5A&
      Caption         =   "2"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00A57A5A&
      Caption         =   "1"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox J1 
      BackColor       =   &H00A78176&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Text            =   "joueur1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox J4 
      BackColor       =   &H00A78176&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Text            =   "joueur4"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox J3 
      BackColor       =   &H00A78176&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "joueur3"
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox J2 
      BackColor       =   &H00A78176&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "joueur2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00A57A5A&
      Caption         =   "4"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00A57A5A&
      Caption         =   "3"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Qui sera l'informaticien le plus rapide du monde ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   56
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Une bonne lettre vaut 3 points"
      Height          =   255
      Left            =   7560
      TabIndex        =   53
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Shape Shape6 
      Height          =   495
      Left            =   8280
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   24.24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   52
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Lettre : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   11.41
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   51
      Top             =   4200
      Width           =   855
   End
   Begin VB.Shape Shape5 
      Height          =   615
      Left            =   9720
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label TpsR3 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.81
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9840
      TabIndex        =   50
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label NbClicks3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   49
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonnes lettres"
      Height          =   255
      Left            =   8280
      TabIndex        =   48
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 1 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   47
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Jeu n°3: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.83
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   46
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Taper sur le clavier les lettres données en 20 secondes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   45
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E2B64E&
      BorderWidth     =   3
      X1              =   7440
      X2              =   7440
      Y1              =   2040
      Y2              =   5880
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 1 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   44
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Clicks"
      Height          =   255
      Left            =   6480
      TabIndex        =   43
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label NbClicks2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   42
      Top             =   3840
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E2B64E&
      BorderWidth     =   4
      X1              =   0
      X2              =   10800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label TpsR2 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.81
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6000
      TabIndex        =   39
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   5880
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Un click vaut 3 points"
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00A78176&
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   2880
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Faire le plus de clicks en 20 secondes sur les boutons rouges"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   33
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jeu n°2: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.83
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   32
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E2B64E&
      BorderWidth     =   3
      X1              =   2760
      X2              =   2760
      Y1              =   2040
      Y2              =   5880
   End
   Begin VB.Label Label11 
      Caption         =   "3"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "pts"
      Height          =   255
      Left            =   1080
      TabIndex        =   29
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label NbPts 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label NbClicks 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Clicks"
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 1 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   1080
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label TpsR 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.81
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1200
      TabIndex        =   23
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Un click vaut 3 points"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Quijoue 
      BackStyle       =   0  'Transparent
      Caption         =   "C'est à joueur1 de jouer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label ptsJ2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label ptsJ3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label PtsJ4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label ptsJ1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   4200
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Faire le plus de clicks en 20 secondes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jeu n°1: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.83
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Prénoms 
      BackStyle       =   0  'Transparent
      Caption         =   "Prénoms :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de joueurs :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label NbJoueurs 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 1 :"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 2 :"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 4 :"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Joueur 3:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Click"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuelTouche As Long
Dim Go As Long
Dim Temps As Long
Dim Temps2 As Long
Dim Temps3 As Long
Dim Tour As Long
Dim Tour2 As Long
Dim Tour3 As Long

Private Sub B1_Click()

Dim ToucheCouleur As Integer
'si le bouton est rouge, on rajoute les points, on remet
'le bouton en gris et on "randomize" un autre bouton
'qui va devenir rouge a son tour
If B1.BackColor = &HFF& Then
    NbClicks2.Caption = NbClicks2.Caption + 1
    If Tour2 = 1 Then ptsJ1.Caption = ptsJ1.Caption + 3
    If Tour2 = 2 Then ptsJ2.Caption = ptsJ2.Caption + 3
    If Tour2 = 3 Then ptsJ3.Caption = ptsJ3.Caption + 3
    If Tour2 = 4 Then PtsJ4.Caption = PtsJ4.Caption + 3
    B1.BackColor = &HA78176
    Randomize
    ToucheCouleur = Int((3 * Rnd) + 1)
        If ToucheCouleur = 1 Then B4.BackColor = &HFF&
        If ToucheCouleur = 2 Then B2.BackColor = &HFF&
        If ToucheCouleur = 3 Then B3.BackColor = &HFF&
    End If
End Sub
Private Sub B2_Click()

Dim ToucheCouleur As Integer
'si le bouton est rouge, on rajoute les points, on remet
'le bouton en gris et on "randomize" un autre bouton
'qui va devenir rouge a son tour
If B2.BackColor = &HFF& Then
    NbClicks2.Caption = NbClicks2.Caption + 1
    If Tour2 = 1 Then ptsJ1.Caption = ptsJ1.Caption + 3
    If Tour2 = 2 Then ptsJ2.Caption = ptsJ2.Caption + 3
    If Tour2 = 3 Then ptsJ3.Caption = ptsJ3.Caption + 3
    If Tour2 = 4 Then PtsJ4.Caption = PtsJ4.Caption + 3
    B2.BackColor = &HA78176
    Randomize
    ToucheCouleur = Int((3 * Rnd) + 1)
        If ToucheCouleur = 1 Then B1.BackColor = &HFF&
        If ToucheCouleur = 2 Then B4.BackColor = &HFF&
        If ToucheCouleur = 3 Then B3.BackColor = &HFF&
    End If
End Sub

Private Sub B3_Click()

Dim ToucheCouleur As Integer
'si le bouton est rouge, on rajoute les points, on remet
'le bouton en gris et on "randomize" un autre bouton
'qui va devenir rouge a son tour
If B3.BackColor = &HFF& Then
    NbClicks2.Caption = NbClicks2.Caption + 1
    If Tour2 = 1 Then ptsJ1.Caption = ptsJ1.Caption + 3
    If Tour2 = 2 Then ptsJ2.Caption = ptsJ2.Caption + 3
    If Tour2 = 3 Then ptsJ3.Caption = ptsJ3.Caption + 3
    If Tour2 = 4 Then PtsJ4.Caption = PtsJ4.Caption + 3
    B3.BackColor = &HA78176
    Randomize
    ToucheCouleur = Int((3 * Rnd) + 1)
        If ToucheCouleur = 1 Then B1.BackColor = &HFF&
        If ToucheCouleur = 2 Then B2.BackColor = &HFF&
        If ToucheCouleur = 3 Then B4.BackColor = &HFF&
    End If
End Sub

Private Sub B4_Click()

Dim ToucheCouleur As Integer
'si le bouton est rouge, on rajoute les points, on remet
'le bouton en gris et on "randomize" un autre bouton
'qui va devenir rouge a son tour
If B4.BackColor = &HFF& Then
    NbClicks2.Caption = NbClicks2.Caption + 1
    If Tour2 = 1 Then ptsJ1.Caption = ptsJ1.Caption + 3
    If Tour2 = 2 Then ptsJ2.Caption = ptsJ2.Caption + 3
    If Tour2 = 3 Then ptsJ3.Caption = ptsJ3.Caption + 3
    If Tour2 = 4 Then PtsJ4.Caption = PtsJ4.Caption + 3
    B4.BackColor = &HA78176
    Randomize
    ToucheCouleur = Int((3 * Rnd) + 1)
        If ToucheCouleur = 1 Then B1.BackColor = &HFF&
        If ToucheCouleur = 2 Then B2.BackColor = &HFF&
        If ToucheCouleur = 3 Then B3.BackColor = &HFF&
    End If
End Sub

Private Sub Clickeur_Click()

'ajout d'un click au comteur par click (normal!!!)
'puis on ajoute 3 points pas click
NbClicks.Caption = NbClicks.Caption + 1
If Tour = 1 Then ptsJ1.Caption = ptsJ1.Caption + 3
If Tour = 2 Then ptsJ2.Caption = ptsJ2.Caption + 3
If Tour = 3 Then ptsJ3.Caption = ptsJ3.Caption + 3
If Tour = 4 Then PtsJ4.Caption = PtsJ4.Caption + 3
End Sub


Private Sub Command1_Click()
'comencer la partie du click

CommenceJ2.Enabled = False
CommenceJ3.Enabled = False
Tour = Tour + 1
Temps = 0
Timer1.Interval = 0
'on passe au joueur suivant
If NbJoueurs.Caption = 1 Then Tour = 1
If NbJoueurs.Caption = 2 And Tour = 3 Then Tour = 1
If NbJoueurs.Caption = 3 And Tour = 4 Then Tour = 1
If NbJoueurs.Caption = 4 And Tour = 5 Then Tour = 1
'on remet le compteur a 0
NbClicks.Caption = 0

Command1.Enabled = False
'affichage du joueur entrain de jouer
If Tour = 1 Then Label10.Caption = J1.Text
If Tour = 2 Then Label10.Caption = J2.Text
If Tour = 3 Then Label10.Caption = J3.Text
If Tour = 4 Then Label10.Caption = J4.Text
'timer se met en route pour compter 20sec ( = 20 * 1 sec)
Timer1.Interval = 1000
Clickeur.Enabled = True
'a qui de jouer ?
Quijoue.Caption = "C'est à " + Label10.Caption + " de jouer."
End Sub

Private Sub Command2_Click()
'on regle l'affichage les label, texts box ... en fontion du nombre de joueurs
NbJoueurs.Caption = "2"
Label5.Visible = True
J2.Visible = True
Label3.Visible = False
J3.Visible = False
Label4.Visible = False
J4.Visible = False

End Sub

Private Sub Command3_Click()
'on regle l'affichage les label, texts box ... en fontion du nombre de joueurs

NbJoueurs.Caption = "3"
Label5.Visible = True
J2.Visible = True
Label3.Visible = True
J3.Visible = True
Label4.Visible = False
J4.Visible = False
End Sub

Private Sub Command4_Click()
'on regle l'affichage les label, texts box ... en fontion du nombre de joueurs

NbJoueurs.Caption = "4"
Label5.Visible = True
J2.Visible = True
Label3.Visible = True
J3.Visible = True
Label4.Visible = True
J4.Visible = True
End Sub

Private Sub Command5_Click()
'on regle l'affichage les label, texts box ... en fontion du nombre de joueurs

NbJoueurs.Caption = "1"
Label5.Visible = False
J2.Visible = False
Label3.Visible = False
J3.Visible = False
Label4.Visible = False
J4.Visible = False
End Sub


Private Sub Command7_Click()
' y a rien a comprendre...
End
End Sub

Private Sub CommenceJ2_Click()
'pour commencer le jeu 2
'initialisation des variables, labels, texts, boutons (le bazar quoi...)
Command1.Enabled = False
CommenceJ3.Enabled = False
Timer3.Interval = 100
Tour2 = Tour2 + 1
Temps2 = 0
Timer2.Interval = 0
'on définit quel joueur va jouer
If NbJoueurs.Caption = 1 Then Tour2 = 1
If NbJoueurs.Caption = 2 And Tour2 = 3 Then Tour2 = 1
If NbJoueurs.Caption = 3 And Tour2 = 4 Then Tour2 = 1
If NbJoueurs.Caption = 4 And Tour2 = 5 Then Tour2 = 1
'on remet a 0 le compteur
NbClicks2.Caption = 0
CommenceJ2.Enabled = False
'affichage du nom du joueur entrain de jouer
If Tour2 = 1 Then Label19.Caption = J1.Text
If Tour2 = 2 Then Label19.Caption = J2.Text
If Tour2 = 3 Then Label19.Caption = J3.Text
If Tour2 = 4 Then Label19.Caption = J4.Text
'c'est parti pour 20 sec ( = 20 * 1 sec)
Timer2.Interval = 1000
'pour pouvoir cliquer sur les boutons
B1.Enabled = True
B2.Enabled = True
B3.Enabled = True
B4.Enabled = True

Quijoue.Caption = "C'est à " + Label19.Caption + " de jouer."

End Sub

Private Sub CommenceJ3_Click()
'pour comencer le jeu no3

CommenceJ3.Enabled = False
CommenceJ2.Enabled = False
Command1.Enabled = False
Tape.Text = "---"
Tour3 = Tour3 + 1
Temps3 = 0
Timer2.Interval = 0
'désigne celui/cele qui joue
If NbJoueurs.Caption = 1 Then Tour3 = 1
If NbJoueurs.Caption = 2 And Tour3 = 3 Then Tour3 = 1
If NbJoueurs.Caption = 3 And Tour3 = 4 Then Tour3 = 1
If NbJoueurs.Caption = 4 And Tour3 = 5 Then Tour3 = 1

'pour pouvoir ecrire dans le textbox et on remet le compteur  0
Tape.Enabled = True
NbClicks3.Caption = 0

'affichage du joueur entrain de jouer
If Tour3 = 1 Then Label20.Caption = J1.Text
If Tour3 = 2 Then Label20.Caption = J2.Text
If Tour3 = 3 Then Label20.Caption = J3.Text
If Tour3 = 4 Then Label20.Caption = J4.Text

'c'est parti pour 20 sec ( = 20 * 1 sec)

Timer5.Interval = 1000
Quijoue.Caption = "C'est à " + Label19.Caption + " de jouer."

'pour définir quelle lettre on va mettre ----> module
DefLettre

End Sub


Private Sub Tape_Change()
'on vérifie si la lettre tapée = lettre donnée
Dim QTape As String
QTape = Right(Tape.Text, 1)
'si c bon, on ajoute les points
If QTape = Label25.Caption Then
    NbClicks3.Caption = NbClicks3.Caption + 1
    If Tour3 = 1 Then ptsJ1.Caption = ptsJ1.Caption + 3
    If Tour3 = 2 Then ptsJ2.Caption = ptsJ2.Caption + 3
    If Tour3 = 3 Then ptsJ3.Caption = ptsJ3.Caption + 3
    If Tour3 = 4 Then PtsJ4.Caption = PtsJ4.Caption + 3
    DefLettre
    End If

End Sub

Private Sub Timer1_Timer()
'pour avoir 20 secondes ...
Timer1.Interval = 1000
Temps = Temps + 1
TpsR.Caption = Temps
If Temps > 20 Then Clickeur.Enabled = False
End Sub

Private Sub Timer2_Timer()
'20 secondes encore et si c'est fini on remet les boutons à leurs place
Timer2.Interval = 1000
Temps2 = Temps2 + 1
TpsR2.Caption = Temps2
If Temps2 > 20 Then
    B1.Enabled = False
    B2.Enabled = False
    B3.Enabled = False
    B4.Enabled = False
    B1.Top = 3720
    B3.Top = 5400
    B2.Top = 3720
    B4.Top = 5400
    B1.Left = 3000
    B3.Left = 3000
    B2.Left = 5400
    B4.Left = 5400
    Timer3.Interval = 0
    Timer4.Interval = 0
    End If
End Sub

Private Sub Timer3_Timer()
'pour savoir comment et ou deplacer quels boutons (touches)
Randomize
Timer4.Interval = Int((500 * Rnd) + 50)

Randomize
QuelTouche = Int((8 * Rnd) + 1)

End Sub

Private Sub Timer4_Timer()

'déplacement des boutons (= touches)
If QuelTouche = 1 Then
    B1.Top = B1.Top + 120
    GoTo 10
    End If
If QuelTouche = 2 Then
    B2.Top = B2.Top + 120
        GoTo 10
    End If
If QuelTouche = 3 Then
    B3.Top = B3.Top + 120
        GoTo 10
    End If
If QuelTouche = 4 Then
    B4.Top = B4.Top + 120
    GoTo 10
    End If
    
If QuelTouche = 5 Then
    B1.Left = B1.Left + 120
    GoTo 10
    End If
If QuelTouche = 6 Then
    B2.Left = B2.Left + 120
    GoTo 10
    End If
If QuelTouche = 7 Then
    B3.Left = B3.Left + 120
    GoTo 10
    End If
If QuelTouche = 8 Then
    B4.Left = B4.Left + 120
    GoTo 10
    End If
    
10

'pour ne pas dépasser les limites du cadre de jeu

If B1.Top > 5520 Then B1.Top = 5000
If B2.Top > 5520 Then B2.Top = 5000
If B3.Top > 5520 Then B3.Top = 5000
If B4.Top > 5520 Then B4.Top = 5000

If B1.Top < 3000 Then B1.Top = 3000
If B2.Top < 3000 Then B2.Top = 3000
If B3.Top < 3000 Then B3.Top = 3000
If B4.Top < 3000 Then B4.Top = 3000


If B1.Left < 3000 Then B1.Left = 3000
If B2.Left < 3000 Then B2.Left = 3000
If B3.Left < 3000 Then B3.Left = 3000
If B4.Left < 3000 Then B4.Left = 3000

If B1.Left > 5400 Then B1.Left = 5200
If B2.Left > 5400 Then B2.Left = 5200
If B3.Left > 5400 Then B3.Left = 5200
If B4.Left > 5400 Then B4.Left = 5200


End Sub

Private Sub Timer5_Timer()
'20 sec pour le jeu de la dactylographie
Timer5.Interval = 1000
Temps3 = Temps3 + 1
TpsR3.Caption = Temps3
If Temps3 > 20 Then Tape.Enabled = False
End Sub

Private Sub TpsR_Change()
'pour réactiver les autres jeux lorsque l'on en a fini un
If TpsR.Caption > 21 Then
    CommenceJ3.Enabled = True
    CommenceJ2.Enabled = True
    Command1.Enabled = True
    TpsR.Caption = "0"
    End If
End Sub

Private Sub TpsR2_Change()
'pour réactiver les autres jeux lorsque l'on en a fini un

If TpsR2.Caption > 21 Then
    CommenceJ3.Enabled = True
    CommenceJ2.Enabled = True
    Command1.Enabled = True
    TpsR2.Caption = "0"
    B1.Enabled = False
    B2.Enabled = False
    B3.Enabled = False
    B4.Enabled = False
    End If
End Sub

Private Sub TpsR3_Change()
'pour réactiver les autres jeux lorsque l'on en a fini un

If TpsR3.Caption > 21 Then
    CommenceJ3.Enabled = True
    CommenceJ2.Enabled = True
    Command1.Enabled = True
    TpsR3.Caption = "0"
    End If
End Sub
