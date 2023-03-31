VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJoueur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_Fond As PictureBox        ' où est dessiné le jeu
Private m_Fond2 As PictureBox       ' pour la prochaine pièce
Private m_TimerJeu As Timer
Private m_Label(3) As Label         ' où sont affichés les résultats

Private m_Touche(4) As Long         ' les 4 touches de mouvements
Private m_bTouche(4) As Boolean     ' état de la touche : pressée ou non

Private m_ChuteLibre As Boolean     ' la pièce chute jusqu'en bas
Private m_DureeChute As Integer     ' vitesse de chute automatique
Private m_Chute As Integer          ' compteur pour la chute
Private m_Animation As Integer      ' compteur pour l'animation

Private m_Grille(MAX_X, MAX_Y) As ECouleur
Private m_GrilleAnim(MAX_X, MAX_Y) As Integer
Private m_Destruction As Integer    ' lignes à ajouter chez l'autre (<=3)

Private m_Piece As CPiece
Private m_ProchainePiece As CPiece

Private m_Niveau As Integer
Private m_Score As Integer
Private m_Lignes As Integer         ' compte les lignes ou les blocs

' Règles pour les points :
' 1 pièce posée = 1                         -> FixePiece
' 1 chute libre = 2                         -> Evenement
' Lignes : 8 * NombreDeLignes^2 + Hauteur   -> Animation
' Blocs  : 8 * (NombreBlocs-3)^2 + Hauteur

Private Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer

' -----------------------------------------------------------------------------
' Nom  : Touche (Property Let)
' -----------------------------------------------------------------------------
Public Property Let Touche(ByVal Index As Integer, ByVal Value As Long)
    m_Touche(Index) = Value
End Property

' -----------------------------------------------------------------------------
' Nom  : Fond / Fond2 / TimerJeu / LabelPoints (Property Set)
' -----------------------------------------------------------------------------
Public Property Set Fond(ByVal Value As PictureBox)
    Set m_Fond = Value
End Property
Public Property Set Fond2(ByVal Value As PictureBox)
    Set m_Fond2 = Value
End Property
Public Property Set TimerJeu(ByVal Value As Timer)
    Set m_TimerJeu = Value
End Property
Public Property Set LabelPoints(ByVal Index As Integer, ByVal Value As Label)
    Set m_Label(Index) = Value
End Property

' -----------------------------------------------------------------------------
' Nom  : ChangeNiveau
' -----------------------------------------------------------------------------
Private Sub ChangeNiveau(ByVal Niveau As Integer)
Dim A As Single, B As Single, C As Single

    If Niveau <= MAX_NIVEAU Then
        m_Niveau = Niveau
        ' vitesse : fonction quadratique du niveau (V = a N^2 + b N + c)
        C = DUREE_CHUTE_INI
        A = (DUREE_CHUTE_INI - DUREE_CHUTE_FIN) / MAX_NIVEAU ^ 2
        B = -2 * A * MAX_NIVEAU
        m_DureeChute = Int(A * Niveau ^ 2 + B * Niveau + C)
        ' changement de tempo
        ChangeTempo Niveau
    End If

End Sub

' -----------------------------------------------------------------------------
' Nom  : SupprimeBloc / SupprimeBlocs
' Desc : Supprime un bloc / les blocs des GrilleAnim
' -----------------------------------------------------------------------------
Private Sub SupprimeBloc(ByVal X As Integer, ByVal Y As Integer)
Dim i As Integer
    For i = Y To 2 Step -1
        m_Grille(X, i) = m_Grille(X, i - 1)
    Next
    m_Grille(X, 1) = VIDE
End Sub
Private Sub SupprimeBlocs()
Dim X As Integer, Y As Integer
    For Y = 1 To MAX_Y
        For X = 1 To MAX_X
            If m_GrilleAnim(X, Y) = -2 Then SupprimeBloc X, Y
        Next
    Next
End Sub

' -----------------------------------------------------------------------------
' Nom  : AjouteLignes
' Desc : Ajoutes N lignes remplies aléatoirement et déplace la pièce
' -----------------------------------------------------------------------------
Public Sub AjouteLignes(ByVal N As Integer)
Dim i As Integer, j As Integer
Dim k As Integer

    ' remplies les lignes
    For j = 1 To MAX_Y
        k = 1
        For i = 1 To MAX_X
            If j <= MAX_Y - N Then
                m_Grille(i, j) = m_Grille(i, j + N)
            Else
                If Rnd > 0.5 And k < MAX_X Then
                    k = k + 1
                    m_Grille(i, j) = Int(Rnd * MAX_COULEUR) + 1
                Else
                    m_Grille(i, j) = VIDE
                End If
            End If
        Next
    Next
    
    ' bouge la pièce vers le haut
    For i = 1 To N
        m_Piece.BougePiece HAUT
        If Not VerifiePiece Then
            m_Piece.BougePiece BAS
            Exit For
        End If
    Next
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : VerifiePiece
' Desc : Vérifie si la pièce se superpose à des blocs de la grille. Si c'est le
'        cas, retourne false
' -----------------------------------------------------------------------------
Private Function VerifiePiece() As Boolean
Dim X As Integer, Y As Integer
Dim i As Integer

    For i = 1 To 4
        m_Piece.Position i, X, Y
        If X < 1 Or X > MAX_X Or Y < 1 Or Y > MAX_Y Then
            VerifiePiece = False: Exit Function
        ElseIf m_Grille(X, Y) > 0 Then
            VerifiePiece = False: Exit Function
        End If
    Next
    VerifiePiece = True
    
End Function

' -----------------------------------------------------------------------------
' Nom  : FixePiece
' Desc : Fixe la pièce dans la grille et vérifie si des lignes sont créées
' -----------------------------------------------------------------------------
Public Sub FixePiece()
Dim X As Integer, Y As Integer, MinY As Integer, MaxY As Integer
Dim i As Integer
Dim bAnimation As Boolean

    bAnimation = False
    m_Destruction = 0
    
    m_Score = m_Score + 1
    
    ' fixe la pièce dans la grille
    MinY = MAX_Y: MaxY = 1
    For i = 1 To 4
        m_Piece.Position i, X, Y
        m_Grille(X, Y) = m_Piece.Couleur(i)
        If Y < MinY Then MinY = Y
        If Y > MaxY Then MaxY = Y
    Next
    
    ' recopie la grille dans GrilleAnim
    ' vérifie directement les lignes entières pour le jeu classique
    For Y = 1 To MAX_Y
        If StyleJeu = JEU_LIGNE Then
            i = 1
        Else
            i = 0
        End If
            
        For X = 1 To MAX_X
            m_GrilleAnim(X, Y) = m_Grille(X, Y)
            If m_Grille(X, Y) = VIDE Then i = 0
        Next
        ' ligne entière
        If i = 1 Then
            bAnimation = True
            m_Destruction = m_Destruction + 1
            For X = 1 To MAX_X
                m_GrilleAnim(X, Y) = -2
            Next
        End If
    Next

    ' destruction de blocs (jeu couleurs) ?
    If StyleJeu = JEU_COULEUR Then
        m_Destruction = VerifieVoisinsIni
        bAnimation = (m_Destruction > 0)
    End If
    
    ' animation
    If bAnimation Then
        m_Animation = 0
        m_TimerJeu.Tag = TAG_ANIM
        m_TimerJeu.Interval = INTERVALLE_ANIM
    Else
        AffichePoints
    End If
        
End Sub

' -----------------------------------------------------------------------------
' Nom  : NouvellePiece
' Rem  : Vérifie si c'est la fin du jeu et dessine la prochaine pièce
' -----------------------------------------------------------------------------
Public Sub NouvellePiece()

    Set m_Piece = m_ProchainePiece
    Set m_ProchainePiece = New CPiece
    DessineProchain
    
    If Not VerifiePiece Then FinJeu
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : DessineProchain
' -----------------------------------------------------------------------------
Public Sub DessineProchain()
Dim X As Integer, XMin As Integer, XMax As Integer, dX As Single
Dim Y As Integer, YMin As Integer, YMax As Integer, dY As Single
Dim i As Integer

    ' dimensions de la prochaine pièce
    XMin = MAX_X: XMax = 0
    YMin = MAX_Y: YMax = 0
    For i = 1 To 4
        m_ProchainePiece.Position i, X, Y
        If X > XMax Then XMax = X
        If X < XMin Then XMin = X
        If Y > YMax Then YMax = Y
        If Y < YMin Then YMin = Y
    Next
    ' position de la prochaine pièce
    dX = (5 - XMax - XMin) / 2
    dY = (5 - YMax - YMin) / 2
    ' dessine la prochaine pièce
    m_Fond2.Cls
    For i = 1 To 4
        m_ProchainePiece.Position i, X, Y
        DessineBloc X + dX, Y + dY, m_ProchainePiece.Couleur(i), m_Fond2, True
    Next

End Sub

' -----------------------------------------------------------------------------
' Nom  : DessineJeu
' -----------------------------------------------------------------------------
Public Sub DessineJeu()
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer

    ' efface tout
    m_Fond.Cls
    
    ' dessine la grille
    For i = 1 To MAX_X
        For j = 1 To MAX_Y
            If m_Grille(i, j) > 0 Then DessineBloc i, j, m_Grille(i, j), m_Fond
        Next
    Next
    
    ' dessine la pièce
    For i = 1 To 4
        m_Piece.Position i, X, Y
        DessineBloc X, Y, m_Piece.Couleur(i), m_Fond
    Next

End Sub

' -----------------------------------------------------------------------------
' Nom  : Animation
' Desc : Fait clignoter les lignes entières et retourne le nombre de lignes
'        moins 1 (handicape pour l'autre)
' -----------------------------------------------------------------------------
Public Function Animation() As Integer
Dim X As Integer, Y As Integer
Dim i As Integer, H As Integer
Dim Couleur As ECouleur
Dim bNiveau As Boolean

    ' dessine les blocs
    Couleur = NOIR
    For X = 1 To MAX_X
        For Y = MAX_Y To MAX_X Step -1
            If m_GrilleAnim(X, Y) = -2 Then
                H = MAX_Y - Y
                If m_Animation Mod 2 = 1 Then Couleur = m_Grille(X, Y)
                DessineBloc X, Y, Couleur, m_Fond
            End If
        Next
    Next
    
    ' fin de l'animation : supprime les blocs
    m_Animation = m_Animation + 1
    If m_Animation > MAX_ANIMATION Then
        SupprimeBlocs
        ' score, niveau
        i = m_Lignes
        m_Lignes = m_Lignes + m_Destruction
        bNiveau = False
        If StyleJeu = JEU_LIGNE Then
            bNiveau = (m_Lignes \ LIGNE_NIVEAU <> i \ LIGNE_NIVEAU)
            m_Score = m_Score + 8 * m_Destruction ^ 2 + H
        Else
            bNiveau = (m_Lignes \ BLOC_NIVEAU <> i \ BLOC_NIVEAU)
            m_Score = m_Score + 8 * (m_Destruction - 3) ^ 2 + H
        End If
        ' changement de niveau
        If bNiveau Then
            PlayWave WAV_NIVEAU
            ChangeNiveau m_Niveau + 1
        Else
            PlayWave WAV_LIGNE
        End If
        ' affiche les points
        AffichePoints
        ' nouvelle pièce et retour au jeu ou fin
        NouvellePiece
        If Not VerifiePiece Then            ' fin du jeu
            FinJeu
        Else
            ' continue avec une nouvelle pièce
            DessineJeu
            m_TimerJeu.Interval = INTERVALLE_JEU
            m_TimerJeu.Tag = TAG_JEU
        End If
        Animation = m_Destruction - 1
    Else
        Animation = 0
    End If
        
        
End Function

' -----------------------------------------------------------------------------
' Nom  : Evenement
' -----------------------------------------------------------------------------
Public Sub Evenement()
Dim i As Integer
Dim bPresse As Boolean, bMouvement As Boolean
Dim bAudio As Boolean
        
    ' vérifie l'état des touches et effectue les mouvements demandés, si
    ' ceux-ci sont possibles
    bMouvement = False
    bAudio = False
    For i = 1 To 4
        bPresse = (GetKeyState(m_Touche(i)) < 0)
        If bPresse And Not m_bTouche(i) Then
            m_Piece.BougePiece i
            If Not VerifiePiece Then                    ' mouvement impossible
                PlayWave WAV_TOC
                m_Piece.BougePiece MouvementInverse(i)
            Else
                bMouvement = True
                If i = BAS And Not m_ChuteLibre Then    ' chute libre
                    m_ChuteLibre = True
                    m_Chute = 0
                    m_Score = m_Score + 2
                ElseIf bAudio = False Then
                    bAudio = True
                    PlayWave WAV_MOUVEMENT
                End If
            End If
        End If
        m_bTouche(i) = bPresse
    Next
    
    ' chute automatique
    m_Chute = m_Chute + 1
    If (m_ChuteLibre And m_Chute >= DUREE_CHUTE_LIBRE) Or _
       (m_Chute >= m_DureeChute) Then
        m_Chute = 0
        bMouvement = True
        m_Piece.BougePiece BAS
        If Not VerifiePiece Then
            m_Piece.BougePiece HAUT
            If m_ChuteLibre Then                        ' fin de la chute libre
                m_ChuteLibre = False
                PlayWave WAV_TOC
            Else                                        ' fixe la pièce
                FixePiece
                ' une animation est en cours
                If m_TimerJeu.Tag = TAG_ANIM Then Exit Sub
                NouvellePiece
                ' fin du jeu
                If m_TimerJeu.Enabled = False Then Exit Sub
            End If
        ' son lorsque la pièce touche le sol
        ElseIf Not m_ChuteLibre Then
            m_Piece.BougePiece BAS
            If Not VerifiePiece Then
                PlayWave WAV_TOC
            End If
            m_Piece.BougePiece HAUT
        End If
    End If
    
    ' la pièce a bougé
    If bMouvement Then DessineJeu
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : AffichePoints
' Desc : Affiche le niveau, le nombre de lignes et le score du joueur
' -----------------------------------------------------------------------------
Private Sub AffichePoints()
    m_Label(1).Caption = CStr(m_Niveau)
    m_Label(2).Caption = CStr(m_Lignes)
    m_Label(3).Caption = CStr(m_Score)
End Sub

' -----------------------------------------------------------------------------
' Nom  : Initialise
' -----------------------------------------------------------------------------
Public Sub Initialise(ByVal Niveau As Integer)
Dim i As Integer, j As Integer
Dim A As Single, B As Single, C As Single

    ' vide la grille
    For i = 1 To MAX_X
        For j = 1 To MAX_Y
            m_Grille(i, j) = VIDE
        Next
    Next
    
    ' touches
    For i = 1 To 4
        m_bTouche(i) = False
    Next
    
    ' crée les premières pièces du jeu
    Set m_ProchainePiece = New CPiece
    Set m_Piece = New CPiece
    
    ' résultats
    m_Lignes = 0
    m_Score = 0
    ChangeNiveau Niveau
    AffichePoints
    
    ' divers
    m_Animation = 0
    m_Chute = 0
    m_ChuteLibre = False
        
End Sub

' -----------------------------------------------------------------------------
' Nom  : VerifieVoisins
' Desc : Regarde si le bloc (X,Y) a des voisins de même couleur
' -----------------------------------------------------------------------------
Private Sub VerifieVoisins(ByVal X As Integer, ByVal Y As Integer, _
    ByVal Couleur As ECouleur, ByRef NbrVoisins As Integer)
    
    If X < 1 Or X > MAX_X Or Y < 1 Or Y > MAX_Y Then    ' en dehors du jeu
        Exit Sub
    ElseIf m_GrilleAnim(X, Y) = Couleur Then            ' bonne couleur
        NbrVoisins = NbrVoisins + 1
        m_GrilleAnim(X, Y) = -1
        VerifieVoisins X + 1, Y, Couleur, NbrVoisins    ' droite
        VerifieVoisins X - 1, Y, Couleur, NbrVoisins    ' gauche
        VerifieVoisins X, Y + 1, Couleur, NbrVoisins    ' bas
        VerifieVoisins X, Y - 1, Couleur, NbrVoisins    ' haut
    End If
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : VerifieVoisinsIni
' Desc : Regarde si plusieurs blocs de même couleur se touchent
' -----------------------------------------------------------------------------
Private Function VerifieVoisinsIni() As Integer
Dim X As Integer, Y As Integer
Dim i As Integer
Dim NbrVoisins As Integer, TotalVoisins As Integer
    
    ' vérifie si les blocs de la pièce ont des voisins de même couleur
    TotalVoisins = 0
    For i = 1 To 4
        NbrVoisins = 0
        m_Piece.Position i, X, Y
        VerifieVoisins X, Y, m_Grille(X, Y), NbrVoisins
        ' prend en compte ou non les voisins trouvés
        For X = 1 To MAX_X
            For Y = 1 To MAX_Y
                If m_GrilleAnim(X, Y) = -1 Then
                     If NbrVoisins >= MIN_VOISINS Then
                        m_GrilleAnim(X, Y) = -2
                     Else
                        m_GrilleAnim(X, Y) = m_Grille(X, Y)
                     End If
                End If
            Next
        Next
        ' nombre total de blocs qui vont être détruits
        If NbrVoisins >= MIN_VOISINS Then
            TotalVoisins = TotalVoisins + NbrVoisins
        End If
    Next
    
    VerifieVoisinsIni = TotalVoisins
    
End Function
