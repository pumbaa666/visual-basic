VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Type TAxe
    X As Integer
    Y As Integer
End Type

Private m_Couleur(4) As ECouleur
Private m_Taille As Integer
Private m_PosRel(4) As TAxe, m_PosAbs As TAxe   ' position relative et absolue

' -----------------------------------------------------------------------------
' Nom  : Couleur (Property Get)
' -----------------------------------------------------------------------------
Public Property Get Couleur(ByVal Index As Integer) As ECouleur
    Couleur = m_Couleur(Index)
End Property

' -----------------------------------------------------------------------------
' Nom  : Position
' Desc : Retourne la position du bloc (Index) dans la grille
' -----------------------------------------------------------------------------
Public Sub Position(ByVal Index As Integer, _
                    ByRef X As Integer, ByRef Y As Integer)
    X = m_PosAbs.X + m_PosRel(Index).X
    Y = m_PosAbs.Y + m_PosRel(Index).Y
End Sub

' -----------------------------------------------------------------------------
' Nom  : BougePiece
' Desc : Effectue une rotation ou une translation de la pièce
' -----------------------------------------------------------------------------
Public Sub BougePiece(ByVal Mouvement As EMouvement)
Dim i As Integer
Dim Temp As Integer

    Select Case Mouvement
        ' translation : change la position absolue (ie dans la grille)
        Case GAUCHE: m_PosAbs.X = m_PosAbs.X - 1
        Case DROITE: m_PosAbs.X = m_PosAbs.X + 1
        Case HAUT: m_PosAbs.Y = m_PosAbs.Y - 1
        Case BAS: m_PosAbs.Y = m_PosAbs.Y + 1
        ' rotation : change les coordonnées relatives des blocs
        Case ROT_POS, ROT_NEG
            For i = 1 To 4
                Temp = m_PosRel(i).X
                If Mouvement = ROT_POS Then
                    m_PosRel(i).X = m_PosRel(i).Y
                    m_PosRel(i).Y = m_Taille - 1 - Temp
                Else
                    m_PosRel(i).X = m_Taille - 1 - m_PosRel(i).Y
                    m_PosRel(i).Y = Temp
                End If
            Next
    End Select
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : Class_Initialize
' Desc : Crée la pièce; positionne les blocs selon un modèle choisi au hasard
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
Dim vPosition As Variant
Dim i As Integer, N As Integer

    ' choisit une des 7 pièces possibles :
    ' 2x2 : 0 = carré
    ' 3x3 : 1 = L, 2 = L inverse, 3 = T, 4 = escalier, 5 = escalier inverse
    ' 4x4 : 6 = I
    N = Int(Rnd * 7)
    Select Case N
        Case 0: vPosition = Array(0, 0, 0, 1, 1, 0, 1, 1): m_Taille = 2
        Case 1: vPosition = Array(1, 0, 1, 1, 1, 2, 2, 2): m_Taille = 3
        Case 2: vPosition = Array(1, 0, 1, 1, 1, 2, 0, 2): m_Taille = 3
        Case 3: vPosition = Array(0, 1, 1, 1, 2, 1, 1, 2): m_Taille = 3
        Case 4: vPosition = Array(2, 0, 1, 1, 2, 1, 1, 2): m_Taille = 3
        Case 5: vPosition = Array(1, 0, 1, 1, 2, 1, 2, 2): m_Taille = 3
        Case 6: vPosition = Array(1, 0, 1, 1, 1, 2, 1, 3): m_Taille = 4
    End Select
    For i = 1 To 4
        m_PosRel(i).X = vPosition(i * 2 - 1)
        m_PosRel(i).Y = vPosition(i * 2)
    Next
    ' effectue entre 0 et 3 rotations, au hasard
    N = Int(Rnd * 4)
    For i = 1 To N
        BougePiece ROT_POS
    Next
    
    ' place la pièce au centre de la grille, en haut
    m_PosAbs.X = 1 + (MAX_X - m_Taille) \ 2
    m_PosAbs.Y = 1
    
    ' couleur
    N = Int(Rnd * MAX_COULEUR) + 1
    For i = 1 To 4
        If StyleJeu = JEU_COULEUR Then N = Int(Rnd * MAX_COULEUR) + 1
        m_Couleur(i) = N
    Next

End Sub
