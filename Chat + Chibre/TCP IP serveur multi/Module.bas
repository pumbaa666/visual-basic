Attribute VB_Name = "Module1"
Public Declare Function cdtInit Lib "Cards.dll" (cWidth As Long, cHeight As Long) As Long
Public Declare Function cdtDrawExt Lib "Cards.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal ordCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long

Public Const PortLocal = 1001
Public Const PortDistant = 1002

Public tTas(3, 8) As Integer
Public tJoueur(3, 3, 8) As Integer
Public tJeu(8, 1) As Integer
Public vCarteEnCours(1) As Integer

Public vLargeur As Long
Public vHauteur As Long

Public Function Distribution()
Dim vCount As Integer
Dim vCount2 As Integer
Dim vCntJoueur As Integer
Dim vCntNumCarte As Integer

Dim vCarte As Byte
Dim vCouleur As Integer
Dim vTmpCouleur As Integer

Dim vNumCarte As Integer
Dim vX As Integer
Dim vY As Integer

Dim vTemp As Integer

'************* Initialisation du tas de carte *************'
    Randomize
    For vCount = 0 To 3
        For vCount2 = 0 To 8
            tTas(vCount, vCount2) = vCount2 + 6
        Next
    Next
'**********************************************************'

'*************** Initialisation des joueurs ***************'
    For vCntJoueur = 0 To 3
        For vCouleur = 0 To 3
            For vCount = 0 To 8
                tJoueur(vCntJoueur, vCouleur, vCount) = 20
            Next
        Next
    Next
'**********************************************************'

'********** Distribution de 9 cartes par joueur ***********'
    For vCntJoueur = 0 To 3
        Do
            vCarte = Int(Rnd * 9) + 6
            vCouleur = Int(Rnd * 4)
            If tTas(vCouleur, vCarte - 6) <> 0 Then
                tJoueur(vCntJoueur, vCouleur, vCntNumCarte) = tTas(vCouleur, vCarte - 6)
                tTas(vCouleur, vCarte - 6) = 0
                vCntNumCarte = vCntNumCarte + 1
            End If
        Loop While (vCntNumCarte < 9)
        vCntNumCarte = 0
    Next
'**********************************************************'

'********************* Tri des cartes *********************'
    For vCntJoueur = 0 To 3
        For vCouleur = 0 To 3
            For vCount = 0 To 8
                For vCntNumCarte = 0 To 7
                    If tJoueur(vCntJoueur, vCouleur, vCntNumCarte) > tJoueur(vCntJoueur, vCouleur, vCntNumCarte + 1) Then
                        vTemp = tJoueur(vCntJoueur, vCouleur, vCntNumCarte)
                        tJoueur(vCntJoueur, vCouleur, vCntNumCarte) = tJoueur(vCntJoueur, vCouleur, vCntNumCarte + 1)
                        tJoueur(vCntJoueur, vCouleur, vCntNumCarte + 1) = vTemp
                    End If
                Next
            Next
        Next
    Next
'**********************************************************'

'******************* Tableau du jeu ***********************'
    vCntNumCarte = 0
    For vTmpCouleur = 0 To 3
        For vCount = 0 To 8
            Select Case vTmpCouleur
                Case 0, 1: vCouleur = vTmpCouleur
                Case 2: vCouleur = 3
                Case 3: vCouleur = 2
            End Select
            
            If tJoueur(0, vCouleur, vCount) <> 20 Then
                tJeu(vCntNumCarte, 0) = tJoueur(0, vCouleur, vCount)
                tJeu(vCntNumCarte, 1) = vCouleur
                vCntNumCarte = vCntNumCarte + 1
            End If
        Next
    Next
'**********************************************************'
End Function

Public Function Affichage()
Dim vCntJoueur As Integer
Dim vCntNumCarte As Integer

Dim vCarte As Byte
Dim vCouleur As Integer
Dim vTmpCouleur As Integer

Dim vNumCarte As Integer
Dim vX As Integer
Dim vY As Integer

'****************** Affichage des cartes ******************'
    Call cdtInit(vLargeur, vHauteur)
    FrmMain.Show
    vCntJoueur = 0
    For vTmpCouleur = 0 To 3
        '************** Pique avant coeur *************'
        Select Case vTmpCouleur
            Case 0, 1: vCouleur = vTmpCouleur
            Case 2: vCouleur = 3
            Case 3: vCouleur = 2
        End Select
        '**********************************************'

        For vCntNumCarte = 0 To 8
'            If tJoueur(vCntJoueur, vCouleur, vCntNumCarte) = 14 Then
            If tJeu(vCntNumCarte, 0) = 14 Then
                vCarte = tJeu(vCntNumCarte, 1)
            Else
'                vCarte = 4 * tJoueur(vCntJoueur, vCouleur, vCntNumCarte) - (4 - vCouleur)
                vCarte = 4 * tJeu(vCntNumCarte, 0) - (4 - tJeu(vCntNumCarte, 1))
            End If

'            If tJoueur(vCntJoueur, vCouleur, vCntNumCarte) = 20 Then
            If tJeu(vCntNumCarte, 0) = 20 Then
                Exit For
            ElseIf tJeu(vCntNumCarte, 0) = 21 Or tJeu(vCntNumCarte, 0) = 22 Then
                vNumCarte = vNumCarte + 1
            Else
                vX = vNumCarte * 35 + 50
                vY = 350
                Call cdtDrawExt(FrmMain.hdc, vX, vY, vLargeur, vHauteur, vCarte, &H0, vbBlue)
                vNumCarte = vNumCarte + 1
            End If
        Next
        vNumCarte = 0
    Next
'**********************************************************'
    FrmMain.Refresh
End Function

