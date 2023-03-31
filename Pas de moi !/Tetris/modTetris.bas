Attribute VB_Name = "modTetris"
Option Explicit
Option Base 1

Public MAX_JOUEUR As Integer
Public Joueur(2) As CJoueur
Public StyleJeu As EStyleJeu

Private NiveauTempo As Integer

' -----------------------------------------------------------------------------
' Nom  : MouvementInverse
' Rem  : Limité aux mouvements que le joueur peut effectués
' -----------------------------------------------------------------------------
Public Function MouvementInverse(ByVal Mouvement As EMouvement) As EMouvement

    Select Case Mouvement
        Case GAUCHE: MouvementInverse = DROITE
        Case DROITE: MouvementInverse = GAUCHE
        Case BAS: MouvementInverse = HAUT
        Case ROT_POS: MouvementInverse = ROT_NEG
    End Select
    
End Function

' -----------------------------------------------------------------------------
' Nom  : InitialiseTetris
' -----------------------------------------------------------------------------
Public Sub InitialiseTetris()
Dim i As Integer, j As Integer
Dim vTouche As Variant

    ' touches pour les mouvements
    vTouche = Array(vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, _
                    vbKeyA, vbKeyD, vbKeyW, vbKeyS)
    
    ' crée les joueurs
    For i = 1 To 2
        Set Joueur(i) = New CJoueur
        With Joueur(i)
            For j = 1 To 4
                .Touche(j) = vTouche(i * 4 + j - 4)
            Next
            Set .Fond = frmMain.pctFond(i)
            Set .Fond2 = frmMain.pctProchain(i)
            Set .TimerJeu = frmMain.tmrJoueur(i)
            For j = 1 To 3
                Set .LabelPoints(j) = frmMain.lblScore(i * 6 + j - 4)
            Next
        End With
    Next
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : ChangeTempo
' Rem  : Tempo en fonction de la vitesse du jeu
' -----------------------------------------------------------------------------
Public Sub ChangeTempo(ByVal Niveau As Integer)
Dim A As Single, B As Single, C As Single
Dim Tempo As Single, Tempo1 As Byte, Tempo2 As Byte

    If Niveau > NiveauTempo Then
        NiveauTempo = Niveau
        ' calcul du tempo en fonction du niveau
        C = TEMPO_INI
        A = (TEMPO_INI - TEMPO_FIN) / MAX_NIVEAU ^ 2
        B = -2 * A * MAX_NIVEAU
        Tempo = A * Niveau ^ 2 + B * Niveau + C
        ' changement dans le fichier midi
        Tempo1 = Int(Tempo)
        Tempo2 = Int((Tempo - Tempo1) * 256)
        TempoMidi Tempo1, Tempo2, 0
        ' joue le fichier midi
        PlayMidi MID_TETRIS
    End If

End Sub

' -----------------------------------------------------------------------------
' Nom  : NouveauJeu
' -----------------------------------------------------------------------------
Public Sub NouveauJeu(ByVal Niveau As Integer)
Dim i As Integer

    ' tempo
    NiveauTempo = -1
    ChangeTempo Niveau
    
    ' jeu
    For i = 1 To MAX_JOUEUR
        Joueur(i).Initialise Niveau
        Joueur(i).DessineJeu
        Joueur(i).DessineProchain
        frmMain.tmrJoueur(i).Tag = TAG_JEU
        frmMain.tmrJoueur(i).Enabled = True
    Next
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : FinJeu
' -----------------------------------------------------------------------------
Public Sub FinJeu()

    frmMain.tmrJoueur(1).Enabled = False
    frmMain.tmrJoueur(2).Enabled = False
    frmMain.mnuJeuPause.Enabled = False
    PlayMidi MID_FIN
    
End Sub
