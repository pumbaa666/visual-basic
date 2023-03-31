VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris"
   ClientHeight    =   12480
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   13440
   Icon            =   "Tetris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   832
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   896
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRepeatMusic 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pctProchain 
      BackColor       =   &H00202040&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   2
      Left            =   480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   15
      Top             =   4320
      Width           =   2415
   End
   Begin VB.PictureBox pctProchain 
      BackColor       =   &H00400040&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   1
      Left            =   10440
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   14
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Timer tmrJoueur 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   10
      Left            =   240
      Top             =   840
   End
   Begin VB.Timer tmrJoueur 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10
      Left            =   12480
      Top             =   840
   End
   Begin VB.PictureBox pctFond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   11700
      Index           =   1
      Left            =   4560
      Picture         =   "Tetris.frx":0CCA
      ScaleHeight     =   778
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   398
      TabIndex        =   0
      Top             =   240
      Width           =   6000
   End
   Begin VB.PictureBox pctFond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   11700
      Index           =   2
      Left            =   2640
      Picture         =   "Tetris.frx":72F3
      ScaleHeight     =   778
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   398
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Math5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Index           =   8
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Niveau"
      BeginProperty Font 
         Name            =   "Math5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Index           =   6
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lignes"
      BeginProperty Font 
         Name            =   "Math5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Index           =   7
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   9
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   10
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   11
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   5
      Left            =   10200
      TabIndex        =   7
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   4
      Left            =   10200
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   3
      Left            =   10200
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lignes"
      BeginProperty Font 
         Name            =   "Math5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Index           =   1
      Left            =   10200
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Niveau"
      BeginProperty Font 
         Name            =   "Math5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Index           =   0
      Left            =   10200
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Math5"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Shape shpContour 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   10
      Height          =   12135
      Index           =   2
      Left            =   240
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Shape shpContour 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   10
      Height          =   12135
      Index           =   1
      Left            =   6600
      Top             =   120
      Width           =   6375
   End
   Begin VB.Menu mnuJeu 
      Caption         =   "&Jeu"
      Begin VB.Menu mnuJeuNouveau 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuJeuPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuJeuBlanc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJeuJoueur 
         Caption         =   "1 Joueur"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuJeuJoueur 
         Caption         =   "2 Joueurs"
         Index           =   2
      End
      Begin VB.Menu mnuJeuBlanc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJeuQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsStyle 
         Caption         =   "Style de jeu"
         Begin VB.Menu mnuStyle 
            Caption         =   "Classique"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuStyle 
            Caption         =   "Spécial Couleurs"
            Index           =   1
         End
      End
      Begin VB.Menu mnuOptionsNiveau 
         Caption         =   "&Niveau initial"
         Begin VB.Menu mnuNiveau 
            Caption         =   "0"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu mnuNiveau 
            Caption         =   "9"
            Index           =   9
         End
      End
      Begin VB.Menu mnuOptionsHandicape 
         Caption         =   "&Handicape"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptionsBlanc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionAudio 
         Caption         =   "&Audio"
         Begin VB.Menu mnuSons 
            Caption         =   "&Sons"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuMusique 
            Caption         =   "&Musique"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private bJeuEnCours As Boolean, bJeuEnPause As Boolean
Private bHandicape As Boolean
Private NiveauIni As Integer
Private AjouteLigne(2) As Integer

' -----------------------------------------------------------------------------
' Nom  : InitialiseJeu
' -----------------------------------------------------------------------------
Private Sub InitialiseJeu()
Dim i As Integer
Dim bJ2 As Boolean

    ' arrête les jeux en cours
    For i = 1 To 2
        tmrJoueur(i).Enabled = False
        tmrJoueur(i).Interval = INTERVALLE_JEU
        pctFond(i).Cls
        pctProchain(i).Cls
    Next
    ' arrête la musique
    StopMidi
    ' affiche ou non le joueur 2
    bJ2 = (MAX_JOUEUR = 2)
    pctFond(2).Visible = bJ2
    pctProchain(2).Visible = bJ2
    shpContour(2).Visible = bJ2
    For i = 6 To 11
        lblScore(i).Visible = bJ2
    Next
    mnuOptionsHandicape.Enabled = bJ2
    ' change le titre des labels en fonction du style de jeu
    If StyleJeu = JEU_COULEUR Then
        lblScore(1).Caption = "Blocs": lblScore(7).Caption = "Blocs"
    Else
        lblScore(1).Caption = "Lignes": lblScore(7).Caption = "Lignes"
    End If
    
    mnuJeuPause.Enabled = False
    bJeuEnCours = False
    bJeuEnPause = False
    
    AjouteLigne(1) = 0
    AjouteLigne(2) = 0
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : Initialise
' -----------------------------------------------------------------------------
Private Sub Initialise()
Dim i As Integer
Dim H As Integer, L1 As Integer, L2 As Integer

    Randomize Timer
    
    ' redimensionne pctFond() en fonction de la taille de l'écran
    With pctFond(2)
        If .Height > 0.8 * Me.ScaleHeight Then .Height = 0.8 * Me.ScaleHeight
        .Width = Int(.Height * MAX_X / MAX_Y)
        .Top = (Me.ScaleHeight - .Height) \ 2
        .Left = Me.ScaleWidth \ 2 - 30 - .Width
        pctFond(1).Height = .Height
        pctFond(1).Width = .Width
        pctFond(1).Top = .Top
        pctFond(1).Left = Me.ScaleWidth \ 2 + 30
    End With
    
    ' graphisme, dimension des blocs, prochaines pièce
    InitialiseGraphique
    With pctProchain(1)
        .Top = pctFond(1).Top
        .Left = (Me.ScaleWidth + pctFond(1).Left + _
                 pctFond(1).Width - .Width) \ 2
        pctProchain(2).Top = .Top
        pctProchain(2).Left = (pctFond(2).Left - .Width) \ 2
    End With
    
    With shpContour(1)
        ' contours des fonds
        .Top = pctFond(1).Top - 10
        .Left = pctFond(1).Left - 10
        .Width = pctFond(1).Width + 20
        .Height = pctFond(1).Height + 20
        shpContour(2).Top = .Top
        shpContour(2).Left = pctFond(2).Left - 10
        shpContour(2).Width = .Width
        shpContour(2).Height = .Height
    End With
    ' labels joueur 1 et 2
    H = pctProchain(1).Top + pctProchain(1).Height + _
        lblScore(0).Height - lblScore(0).Top
    L1 = (Me.ScaleWidth + pctFond(1).Left + _
         pctFond(1).Width - lblScore(0).Width) \ 2
    L2 = (pctFond(2).Left - lblScore(6).Width) \ 2
    For i = 0 To 5
        lblScore(i).Left = L1
        lblScore(i).Top = lblScore(i).Top + H
        lblScore(i + 6).Left = L2
        lblScore(i + 6).Top = lblScore(i).Top
    Next
    
    MAX_JOUEUR = 1
    NiveauIni = 0
    bHandicape = True
    StyleJeu = JEU_LIGNE
    
    ' autres initialisations
    InitialiseTetris
    InitialiseJeu
    InitialiseAudio
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : ChangePause
' -----------------------------------------------------------------------------
Private Sub ChangePause(ByVal Pause As Boolean)

    If Not bJeuEnCours Then Exit Sub
    bJeuEnPause = Pause
    
    If Pause Then
        tmrJoueur(1).Enabled = False
        tmrJoueur(2).Enabled = False
        StopMidi
    Else
        tmrJoueur(1).Enabled = True
        tmrJoueur(2).Enabled = (MAX_JOUEUR = 2)
        If bMidi Then PlayMidi MID_TETRIS
    End If
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : Form_Load / Form_QueryUnload
' -----------------------------------------------------------------------------
Private Sub Form_Load()
    Me.Show
    Initialise
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    StopMidi True
End Sub

' -----------------------------------------------------------------------------
' Nom  : Form_Resize
' Desc : Arrête le jeu quand la fenêtre est fermée, et le relance si elle
'        est restaurée
' -----------------------------------------------------------------------------
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ChangePause True
    ElseIf Me.WindowState = vbMaximized Then
        ChangePause True
    End If
End Sub

' -----------------------------------------------------------------------------
' Nom  : mnuJeuNouveau / mnuJeuPause / mnuJeuJoueur / mnuJeuQuitter (Click)
' Desc : Fonctions du menu Jeu
' -----------------------------------------------------------------------------
Private Sub mnuJeuNouveau_Click()
    bJeuEnCours = True
    mnuJeuPause.Enabled = True
    NouveauJeu NiveauIni
    PlayMidi MID_TETRIS
End Sub
Private Sub mnuJeuPause_Click()
    mnuJeuPause.Checked = Not mnuJeuPause.Checked
    ChangePause mnuJeuPause.Checked
End Sub
Private Sub mnuJeuJoueur_Click(Index As Integer)
    mnuJeuJoueur(Index).Checked = True
    mnuJeuJoueur(3 - Index).Checked = False
    MAX_JOUEUR = Index
    InitialiseJeu
End Sub
Private Sub mnuJeuQuitter_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' Nom  : mnuStyle / mnuNiveau / mnuOptionsHandicape / mnuSons
'        / mnuMusique (Click)
' Desc : Fonctions du menu Options
' -----------------------------------------------------------------------------
Private Sub mnuStyle_Click(Index As Integer)
    mnuStyle(Index).Checked = True
    mnuStyle(1 - Index).Checked = False
    If Index = 0 Then StyleJeu = JEU_LIGNE
    If Index = 1 Then StyleJeu = JEU_COULEUR
    InitialiseJeu
End Sub
Private Sub mnuNiveau_Click(Index As Integer)
    mnuNiveau(NiveauIni).Checked = False
    NiveauIni = Index
    mnuNiveau(NiveauIni).Checked = True
End Sub
Private Sub mnuOptionsHandicape_Click()
    mnuOptionsHandicape.Checked = Not mnuOptionsHandicape.Checked
    bHandicape = mnuOptionsHandicape.Checked
End Sub
Private Sub mnuSons_Click()
    mnuSons.Checked = Not mnuSons.Checked
    bWave = mnuSons.Checked
End Sub
Private Sub mnuMusique_Click()
    mnuMusique.Checked = Not mnuMusique.Checked
    bMidi = mnuMusique.Checked
    If bMidi And bJeuEnCours Then PlayMidi MID_TETRIS
    If Not bMidi Then StopMidi
End Sub

' -----------------------------------------------------------------------------
' Nom  : tmrJoueur / tmrRepeatMusic (Timer)
' -----------------------------------------------------------------------------
Private Sub tmrJoueur_Timer(Index As Integer)
Dim Handicape As Integer

    If tmrJoueur(Index).Tag = TAG_JEU Then
        ' ajoute des lignes si nécessaire
        If AjouteLigne(Index) > 0 Then
            Joueur(Index).AjouteLignes AjouteLigne(Index)
            AjouteLigne(Index) = 0
        End If
        ' jeu du joueur
        Joueur(Index).Evenement
    Else
        ' animation
        Handicape = Joueur(Index).Animation
        ' ajoute des lignes à l'autre (lorsqu'il est en jeu, pas en animation)
        If Handicape > 0 And bHandicape And MAX_JOUEUR = 2 Then
            AjouteLigne(3 - Index) = Handicape
        End If
    End If
        
End Sub
Private Sub tmrRepeatMusic_Timer()
    If bJeuEnCours Then RepeatMidi
End Sub
