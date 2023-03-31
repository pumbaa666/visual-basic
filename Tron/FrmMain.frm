VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Tron"
   ClientHeight    =   4755
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkAveugle 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1200
      Top             =   480
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Timer ClkMain 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   120
         Top             =   2760
      End
      Begin VB.CommandButton CmdQuitter 
         Caption         =   "&Quitter"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton CmdStart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdMulti 
         Caption         =   "&Options Multi-joueur"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Timer ClkQuit 
         Enabled         =   0   'False
         Interval        =   150
         Left            =   1800
         Top             =   3960
      End
      Begin MSComDlg.CommonDialog Couleur 
         Left            =   600
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Changement :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LblIP 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label LblTemps 
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Temps :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Score : "
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label LblScore 
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label LblChange 
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.Shape ShpAveugle 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Shape ShpBords 
      Height          =   135
      Left            =   120
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape ShpTerrain 
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   135
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu FichierStart 
         Caption         =   "Démarrer"
      End
      Begin VB.Menu Tiret1 
         Caption         =   "-"
      End
      Begin VB.Menu FichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu Option 
      Caption         =   "Option"
      Begin VB.Menu OptionPartie 
         Caption         =   "Partie"
         Begin VB.Menu OptionPartieSolo 
            Caption         =   "Solo"
            Checked         =   -1  'True
         End
         Begin VB.Menu OptionPartieMulti 
            Caption         =   "Multi-joueur"
         End
      End
      Begin VB.Menu Tiret2 
         Caption         =   "-"
      End
      Begin VB.Menu OptionCouleur 
         Caption         =   "Choisir sa couleur"
      End
      Begin VB.Menu OptionGrillage 
         Caption         =   "Masquer le grillage"
      End
      Begin VB.Menu OptionChat 
         Caption         =   "Afficher le chat"
      End
   End
   Begin VB.Menu Difficulte 
      Caption         =   "Difficulté"
      Begin VB.Menu DifficulteFacile 
         Caption         =   "Facile"
      End
      Begin VB.Menu DifficulteMoyen 
         Caption         =   "Moyen"
         Checked         =   -1  'True
      End
      Begin VB.Menu DifficulteDifficile 
         Caption         =   "Difficile"
      End
      Begin VB.Menu Tiret4 
         Caption         =   "-"
      End
      Begin VB.Menu DifficulteObstacles 
         Caption         =   "Ajouter des obstacles"
      End
      Begin VB.Menu Tiret5 
         Caption         =   "-"
      End
      Begin VB.Menu DifficulteRandom 
         Caption         =   "Mode ""aléatoire"""
      End
      Begin VB.Menu DifficulteFrustrant 
         Caption         =   "Mode ""frustrant"""
      End
      Begin VB.Menu DifficulteAveugle 
         Caption         =   "Mode ""aveugle"""
      End
      Begin VB.Menu DifficulteParkinson 
         Caption         =   "Mode ""parkinson"""
      End
   End
   Begin VB.Menu Aide 
      Caption         =   "Aide"
      Begin VB.Menu AideJouer 
         Caption         =   "Comment jouer"
      End
      Begin VB.Menu AideScore 
         Caption         =   "Scores"
      End
      Begin VB.Menu AideDifficulte 
         Caption         =   "Difficultés"
      End
      Begin VB.Menu AideMulti 
         Caption         =   "Initialiser une partie Multi"
      End
      Begin VB.Menu Tiret3 
         Caption         =   "-"
      End
      Begin VB.Menu AideAbout 
         Caption         =   "A propos"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTerrain As Boolean
Dim vDif As Integer
Dim vNbChange As Integer
Dim vTempDX As Integer
Dim vTempDY As Integer
Dim vFlushBuffer As Boolean
Dim vDifObstacles As Integer

Private Sub AideAbout_Click()
    FrmAbout.Show
End Sub

Private Sub AideDifficulte_Click()
    FrmAideDif.Show
End Sub

Private Sub AideJouer_Click()
    FrmAideJeu.Show
End Sub

Private Sub AideMulti_Click()
    FrmAideMulti.Show
End Sub

Private Sub AideScore_Click()
    FrmAideScore.Show
End Sub

Private Sub ClkAveugle_Timer()
Static vCntAveugle As Boolean
    If vCntAveugle = False Then
        If Int(Rnd * 7 - vDif) = 0 Then
            vCntAveugle = True
            ShpAveugle.Visible = True
        End If
    Else
        ShpAveugle.Visible = False
        vCntAveugle = False
    End If
End Sub

Private Sub ClkMain_Timer()
Dim tSendCoo(1) As String
Dim vRandom As Integer
Static vTemps As Integer
Dim vMultiple As Integer
Dim vTempDif As Integer

    vFlushBuffer = True
'************************* Temps **************************'
    If vTemps = 20 Then
        LblTemps.Caption = CDate(LblTemps.Caption) + CDate("00:00:01")
        vTemps = 0
    End If
    vTemps = vTemps + 1
'**********************************************************'

'********************* Mode aléatoire *********************'
    On Error Resume Next
    If DifficulteRandom.Checked = True Then
        vRandom = Int(Rnd * (50 - vDif * 5))
        If vRandom = 0 Then
            vNbChange = vNbChange + 1
            If vDX <> 0 Then
                vDX = 0
                vDY = 1
            Else
               vDX = 1
               vDY = 0
            End If
        ElseIf vRandom = 1 Then
            vNbChange = vNbChange + 1
            If vDX <> 0 Then
                vDX = 0
                vDY = -1
            Else
               vDX = -1
               vDY = 0
            End If
        End If
    End If
'**********************************************************'
    
'********************* Mode Parkinson *********************'
    If DifficulteParkinson.Checked = True Then
        FrmMain.Top = FrmMain.Top + (Int(Rnd * 50) - 25) * vDif
        FrmMain.Left = FrmMain.Left + (Int(Rnd * 50) - 25) * vDif
    End If
'**********************************************************'
    tCoo(X) = tCoo(X) + vDX
    tCoo(Y) = tCoo(Y) - vDY

    If tCoo(X) = -1 Or tCoo(Y) = -1 Or tCoo(X) = (DimX + 1) Or tCoo(Y) = DimY + 1 Or tTerrain(tCoo(X), tCoo(Y)) = 1 Then
        vMultiple = 1
        If DifficulteObstacles.Checked = True Then
            vMultiple = 4
        End If
        If DifficulteRandom.Checked = True Then
            vMultiple = vMultiple + 4
        End If
        If DifficulteAveugle.Checked = True Then
            vMultiple = vMultiple + 4
        End If
        If DifficulteParkinson.Checked = True Then
            vMultiple = vMultiple + 4
        End If
        If DifficulteFrustrant.Checked = True Then
            vMultiple = vMultiple + 4
        End If
        vTempDif = vDif
        If vMultiple = 1 Then vTempDif = 1
        LblScore.Caption = (60 * (Mid(LblTemps.Caption, 4, 2)) + Int(Right(LblTemps.Caption, 2))) * 100 * vTempDif * vMultiple * (vNbChange / 20)
        ClkMain.Enabled = False
        FrmOptMulti.Wsk.SendData "[PERDU]"
        OptionCouleur.Enabled = True
        Picture1.Visible = False
        OptionPartie.Enabled = True
        FichierStart.Enabled = True
        CmdStart.Enabled = True
        CmdStart.SetFocus
        vTempDX = 1
        vTempDY = 0
        ClkAveugle.Enabled = False
        MsgBox "Vous avez perdu", vbCritical, "Game Over"
    Else
        If tCoo(X) < 10 Then
            tSendCoo(X) = "0" & tCoo(X)
        Else
            tSendCoo(X) = tCoo(X)
        End If
        If tCoo(Y) < 10 Then
            tSendCoo(Y) = "0" & tCoo(Y)
        Else
            tSendCoo(Y) = tCoo(Y)
        End If
        If FrmOptMulti.Wsk.State = 7 Then
            FrmOptMulti.Wsk.SendData "[COO]" & tSendCoo(X) & ";" & tSendCoo(Y)
        End If
        tTerrain(tCoo(X), tCoo(Y)) = 1
        ShpTerrain(tCoo(Y) * (DimX + 1) + tCoo(X)).FillColor = vCouleur
    End If
End Sub

Private Sub ClkQuit_Timer()
Static vCntQuit As Boolean

    If vCntQuit = False Then
        If FrmOptMulti.Wsk.State = 7 Then
            FrmOptMulti.Wsk.SendData "[QUIT]"
        End If
        vCntQuit = True
    Else
        ClkQuit.Enabled = False
        FrmOptMulti.Wsk.Close
        End
    End If
End Sub

Private Sub CmdMulti_Click()
    OptionPartieMulti_Click
End Sub

Private Sub CmdQuitter_Click()
    ClkQuit.Enabled = True
End Sub

Private Sub CmdStart_Click()
    If vTerrain = False Then
        ClearTerrain
        If DifficulteObstacles.Checked = True Then
            AddObstacles
        End If
    ElseIf vDifObstacles <> vDif And DifficulteObstacles.Checked = True Then
        AddObstacles
    End If

    If OptionPartieMulti.Checked = True Then
        FrmOptMulti.Wsk.SendData "[START]"
        Start
    Else
        vDY = 0
        vDX = 1
        tCoo(X) = 0
        tCoo(Y) = 0
        tTerrain(0, 0) = 1
        ShpTerrain(0).FillColor = vCouleur
        OptionPartieSolo.Checked = True
        OptionPartieMulti.Checked = False
        Picture1.Visible = True
    End If
    
    If DifficulteAveugle.Checked = True Then
        ShpAveugle.Left = ShpTerrain(0).Left
        ShpAveugle.Top = ShpTerrain(0).Top
        ShpAveugle.Width = ShpTerrain(0).Width * (DimX - 4) - ShpTerrain(0).Left
        ShpAveugle.Height = ShpTerrain(0).Height * (DimY - 2) - ShpTerrain(0).Top
        ClkAveugle.Enabled = True
    End If
    
    LblScore.Caption = "0"
    vNbChange = 1
    LblTemps.Caption = "00:00:00"
    FichierStart.Enabled = False
    Picture1.Visible = True
    Picture1.SetFocus
    CmdStart.Enabled = False
    ClkMain.Enabled = True
    OptionPartie.Enabled = False
    OptionCouleur.Enabled = False
    vTerrain = False
End Sub

Private Sub DifficulteAveugle_Click()
    DifficulteAveugle.Checked = Not DifficulteAveugle.Checked
End Sub

Private Sub DifficulteFacile_Click()
    DifficulteFacile.Checked = True
    DifficulteMoyen.Checked = False
    DifficulteDifficile.Checked = False
    vDif = 1
End Sub

Private Sub DifficulteFrustrant_Click()
    DifficulteFrustrant.Checked = Not DifficulteFrustrant.Checked
End Sub

Private Sub DifficulteMoyen_Click()
    DifficulteFacile.Checked = False
    DifficulteMoyen.Checked = True
    DifficulteDifficile.Checked = False
    vDif = 2
End Sub

Private Sub DifficulteDifficile_Click()
    DifficulteFacile.Checked = False
    DifficulteMoyen.Checked = False
    DifficulteDifficile.Checked = True
    vDif = 3
End Sub

Private Sub DifficulteObstacles_Click()
    If DifficulteObstacles.Checked = True Then
        DifficulteObstacles.Checked = False
    Else
        DifficulteObstacles.Checked = True
        AddObstacles
    End If
End Sub

Private Sub DifficulteParkinson_Click()
    DifficulteParkinson.Checked = Not DifficulteParkinson.Checked
End Sub

Private Sub DifficulteRandom_Click()
    DifficulteRandom.Checked = Not DifficulteRandom.Checked
End Sub

Private Sub FichierQuitter_Click()
    If FrmOptMulti.Wsk.State = 7 Then
        FrmOptMulti.Wsk.SendData "[QUIT]"
        FrmOptMulti.Wsk.Close
    End If
    End
End Sub

Private Sub FichierStart_Click()
    CmdStart_Click
End Sub

Private Sub Form_Load()
Dim vX As Integer
Dim vY As Integer

    LblIP.Caption = "Votre adresse IP : " & FrmOptMulti.Wsk.LocalIP
    vDif = 2

    For vY = 0 To DimY
        For vX = 0 To DimX
            If vX + vY <> 0 Then
                Load ShpTerrain(vY * (DimX + 1) + vX)
                ShpTerrain(vY * (DimX + 1) + vX).Top = vY * 120 + 120
                ShpTerrain(vY * (DimX + 1) + vX).Left = vX * 120 + 120
                ShpTerrain(vY * (DimX + 1) + vX).Visible = True
            End If
        Next
    Next

    Frame.Left = DimX * 120 + 400
    FrmMain.Width = Frame.Left + Frame.Width + 200
    FrmMain.Height = DimY * 120 + 200
    If FrmMain.Height < 5500 Then
        FrmMain.Height = 5500
    End If

    vTerrain = True
    vFlushBuffer = True
    ShpBords.Width = DimX * 120 + 130
    ShpBords.Height = DimY * 120 + 130

    On Error GoTo CreatFile
    Open "c:\tron.ini" For Input As #1
    Close #1
    vCouleur = &H0
    Exit Sub

CreatFile:
    Close #1
    Open "c:\tron.ini" For Output As #1
    Print #1, "&H0"
    Close #1
    vCouleur = &H0
End Sub

Private Sub OptionChat_Click()
    If FrmChat.Visible = False Then
        FrmChat.Show
    End If
End Sub

Private Sub OptionCouleur_Click()
    Couleur.ShowColor
    vCouleur = Couleur.Color
    If OptionPartieSolo.Checked = False Then
        FrmOptMulti.Wsk.SendData "[COULEUR]" & vCouleur
    End If
    If vTerrain = False Then
        ClearTerrain
    End If
End Sub

Private Sub OptionGrillage_Click()
Dim vCount As Integer

    If OptionGrillage.Checked = False Then
        For vCount = 0 To (DimX + 1) * (DimY + 1) - 1
            ShpTerrain(vCount).BorderColor = &H8000000F
        Next
        OptionGrillage.Checked = True
    Else
        For vCount = 0 To (DimX + 1) * (DimY + 1) - 1
            ShpTerrain(vCount).BorderColor = &H0
        Next
        OptionGrillage.Checked = False
    End If
End Sub

Private Sub OptionPartieMulti_Click()
    FrmOptMulti.Show
    OptionPartieSolo.Checked = False
    OptionPartieMulti.Checked = True
    CmdStart.Enabled = False
    CmdMulti.Enabled = True
    FichierStart.Enabled = False
End Sub

Private Sub OptionPartieSolo_Click()
    OptionPartieSolo.Checked = True
    OptionPartieMulti.Checked = False
    CmdStart.Enabled = True
    CmdMulti.Enabled = False
    FichierStart.Enabled = True
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Const HAUT = 40
Const BAS = 38
Const GAUCHE = 37
Const DROITE = 39

    If vFlushBuffer = True Then
        vTempDX = vDX
        vTempDY = vDY
        
        vFlushBuffer = False
        
        If ClkMain.Enabled = False Then ClkMain.Enabled = True
        
        If Not (DifficulteFrustrant.Checked = True And Int(Rnd * 7) - vDif = 0) Then
            If KeyCode = HAUT And vDY = 0 Then
                vDY = -1
                vDX = 0
                vNbChange = vNbChange + 1
            ElseIf KeyCode = BAS And vDY = 0 Then
                vDY = 1
                vDX = 0
                vNbChange = vNbChange + 1
            ElseIf KeyCode = GAUCHE And vDX = 0 Then
                vDX = -1
                vDY = 0
                vNbChange = vNbChange + 1
            ElseIf KeyCode = DROITE And vDX = 0 Then
                vDX = 1
                vDY = 0
                vNbChange = vNbChange + 1
            ElseIf KeyCode = 13 And OptionPartieMulti.Checked = False Then
                ClkMain.Enabled = False
                vFlushBuffer = True
                vNbChange = vNbChange - 1
            End If
            LblChange.Caption = vNbChange
        End If
    End If
End Sub

Function AddObstacles()
Dim vCntDif As Integer
Dim vX As Integer
Dim vY As Integer
Dim vColor As Integer

    If vCouleur = 255 Then
        vColor = 0
    Else
        vColor = 255
    End If

    ClearTerrain
    Randomize
    For vCntDif = 0 To (DimX - DimX / 10) * vDif
        vX = Int(Rnd * (DimX + 1))
        vY = Int(Rnd * DimY) + 1
        tTerrain(vX, vY) = 1
        ShpTerrain(vY * (DimX + 1) + vX).FillColor = vColor
    Next
    vTerrain = True
    vDifObstacles = vDif
End Function
