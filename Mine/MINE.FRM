VERSION 5.00
Begin VB.Form FrmJeu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Démineur"
   ClientHeight    =   2895
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCache 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"MINE.frx":0000
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   0
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   2160
      TabIndex        =   64
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   64
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   63
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   62
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   61
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   60
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   59
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   58
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   57
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   56
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   55
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   54
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   53
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   52
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   51
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   50
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   49
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   48
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   47
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   46
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   45
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   44
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   43
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   42
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   41
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   40
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   39
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   38
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   37
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   36
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   35
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   34
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   33
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   32
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   31
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   30
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   29
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   28
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   27
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   26
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   25
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   24
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   23
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   22
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   21
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   20
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   19
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   18
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   17
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   16
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   15
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   14
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   13
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   12
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   11
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   10
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   9
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   8
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   7
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   5
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   4
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   255
   End
   Begin VB.Label LblMine 
      Caption         =   "Mines restantes : 10"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LblTemps 
      Caption         =   "Temps écoulé : 0 sec"
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu MenuNew 
         Caption         =   "Nouveau"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuScore 
         Caption         =   "Meilleur scores"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuQuit 
         Caption         =   "Quitter"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FrmJeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTemps As Integer
Dim vNbMine As Integer
Dim vTestBleu As Integer
Dim TabMine(9, 9) As Integer

Private Sub CmdCache_Click()
    Timer1.Enabled = True
    CmdCache.Visible = False
End Sub

Private Sub CmdPause_Click()
    Timer1.Enabled = False
    CmdCache.Visible = True
End Sub

Private Sub Command1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 97 Then
        If Command1(Index).Caption = "" Then
            If Command1(Index).BackColor <> &HFF0000 Then
                Command1(Index).BackColor = &HFF0000
                vNbMine = vNbMine - 1
            Else
                Command1(Index).BackColor = &HC0C0C0
                vNbMine = vNbMine + 1
                vTestBleu = 1
            End If
            LblMine.Caption = "Mines restantes : " & Str(vNbMine)
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim vX As Integer
Dim vY As Integer
Dim vCount As Integer
Dim vCount2 As Integer
Dim vMineRound As Integer
Dim vGagne As Integer
Dim vNom As String
Dim vTempsEnr As Integer
Dim vTestFile As Boolean
    If Command1(Index).BackColor <> &HFF0000 And vTestBleu <> 1 Then
'   N'affiche pas le carré du focus
        CmdPause.Enabled = True
        CmdPause.SetFocus
'   Démarre la partie si ce n'est pas déjà fait
        If vTemps = 0 Then
            Timer1.Enabled = True
        End If
'   Test si il y a une mine là où on a cliqué
        For vCount = 1 To 8
            If Index < (8 * vCount) + 1 Then
                vY = vCount - 1
                vX = (Index - vY * 8) - 1
                vCount = 9
            End If
        Next
'   Quitte si on perd
        If TabMine(vX, vY) = 1 Then
    '   Affiche toutes les mines
            For vCount = 0 To 8
                For vCount2 = 0 To 8
                    If (vCount2 < 8 Or vCount < 8) Then
                        If TabMine(vCount, vCount2) = 1 Then
                            'Command1((vCount2 * 8 + vCount) + 1).BackColor = &HFF
                            'Command1((vCount2 * 8 + vCount) + 1).Caption = "X"
                            Command1((vCount2 * 8 + vCount) + 1).Picture = LoadPicture("mine.bmp")
                        End If
                    End If
                Next
            Next
            MsgBox "Désolé, vous avez perdu!", vbOKOnly, "Perdu"
'       Désactive les boutons
            For vCount = 0 To 8
                For vCount2 = 0 To 8
                    If (vCount2 < 8 Or vCount < 8) Then
                        Command1((vCount2 * 7 + vCount) + 1).Enabled = False
                    End If
                Next
            Next
            CmdPause.Enabled = False
            Timer1.Enabled = False
        End If
'   Affiche le nombre de mines alentours
        For vCount = vX - 1 To vX + 1
            If vCount = -1 Then
                vCount = 0
            End If
            For vCount2 = vY - 1 To vY + 1
                If vCount2 = -1 Then
                    vCount2 = 0
                End If
                If TabMine(vCount, vCount2) = 1 Then
                    vMineRound = vMineRound + 1
                End If
            Next
        Next
    
'   N'affiche pas le nb de mine alentour si il y a une mine où on a cliqué
        For vCount = 1 To 8
            If Index < (8 * vCount) + 1 Then
                vY = vCount - 1
                vX = (Index - vY * 8) - 1
                vCount = 9
            End If
        Next
        If TabMine(vX, vY) <> 1 Then
            Command1(Index).Caption = Str(vMineRound)
        End If
    
'   Compte le nb de boutons non-cliqués pour savoir si on a gagné
        vGagne = 54
        For vCount = 1 To 64
            If Command1(vCount).Caption <> "" Then
                vGagne = vGagne - 1
            End If
        Next
        If vGagne = 0 Then
            MsgBox "Félicitation, vous avez gagné", vbOKOnly, "Gagné"
            Do
                Timer1.Enabled = False
            Loop While (Timer1.Enabled = True)
'       Désactive les boutons
            For vCount = 0 To 8
                For vCount2 = 0 To 8
                    If (vCount2 < 8 Or vCount < 8) Then
                        Command1((vCount2 * 7 + vCount) + 1).Enabled = False
                    End If
                Next
            Next
            CmdPause.Enabled = False
            Timer1.Enabled = False
'       Crée le fichier si il n'existe pas
            On Error GoTo CreatFile
            Open "c:\temp\scoremine.dat" For Input As #1
            Close #1
            vTestFile = 1
'       lit les score
            Open "c:\temp\scoremine.dat" For Input As #1
            Line Input #1, vNom
            vTempsEnr = Right(vNom, 2)
            Close #1
            If vTemps < vTempsEnr Then
                vNom = InputBox("Vous êtes le meilleur! Entrez votre nom", "Highscore")
                Open "c:\temp\scoremine.dat" For Output As #1
                Print #1, vNom & Str(vTemps)
                Close #1
            Else: MsgBox "Désolé, vous n'êtes pas spécialement mauvais, mais vous n'êtes pas assez rapide!", vbOKOnly, "Sorry"
            End If
            vTemps = 0
CreatFile:
            If vTestFile = 0 Then
                Open "c:\temp\scoremine.dat" For Append As #1
                vNom = InputBox("Vous êtes le meilleur! Entrez votre nom", "Highscore")
                Print #1, vNom & Str(vTemps)
                Close #1
            End If
        End If
    ElseIf vTestBleu = 1 Then
        vTestBleu = 0
    End If
End Sub

Private Sub Form_Load()
Dim vCount As Integer
Dim vCount2 As Integer
Dim vRandx As Integer
Dim vRandy As Integer
Dim vChaine As String
    Randomize
    vNbMine = 0
    vTemps = 0
    LblTemps.Caption = "Temps écoulé : 0 sec"
    Timer1.Enabled = False
    For vCount = 0 To 8
         For vCount2 = 0 To 8
             If (vCount2 < 8 Or vCount < 8) Then
                 TabMine(vCount, vCount2) = 0
                 Command1((vCount2 * 7 + vCount) + 1).Picture = LoadPicture()
                 Command1((vCount2 * 7 + vCount) + 1).Caption = ""
                 Command1((vCount2 * 7 + vCount) + 1).BackColor = &HC0C0C0
                 Command1((vCount2 * 7 + vCount) + 1).Enabled = True
             End If
         Next
    Next
    Do
        vRandx = Int(Rnd * 8)
        vRandy = Int(Rnd * 8)
        If TabMine(vRandx, vRandy) = 0 Then
            TabMine(vRandx, vRandy) = 1
'   TRICHE    TRICHE    TRICHE    TRICHE    TRICHE
'            Command1((vRandy * 8 + vRandx) + 1).BackColor = &HFF
'            Command1((vRandy * 8 + vRandx) + 1).Enabled = False
            vNbMine = vNbMine + 1
        End If
    Loop Until (vNbMine = 10)
    LblMine.Caption = "Mines restantes: " + Str(vNbMine)
End Sub

Private Sub MenuNew_Click()
    Form_Load
End Sub

Private Sub MenuQuit_Click()
    End
End Sub

Private Sub MenuScore_Click()
    Timer1.Enabled = False
    CmdCache.Visible = True
    FrmJeu.Hide
    FrmScore.Show
End Sub

Private Sub Timer1_Timer()
    vTemps = vTemps + 1
    LblTemps.Caption = "Temps écoulé :" + Str(vTemps) + " sec"
End Sub
