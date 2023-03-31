Attribute VB_Name = "modGraphique"
Option Explicit
Option Base 1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private L As Integer                        ' côté d'un bloc
Private D As Integer                        ' épaisseur de la bande de relief
Private Ox As Integer, Oy As Integer        ' origine des axes
Private C1 As Long, C2 As Long, C3 As Long  ' 3 couleurs par bloc

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, _
    lpPoint As POINTAPI, ByVal nCount As Long) As Long

' -----------------------------------------------------------------------------
' Nom  : InitialiseGraphique
' -----------------------------------------------------------------------------
Public Sub InitialiseGraphique()

    ' dimension et position des blocs
    With frmMain.pctFond(1)
        L = .Width \ MAX_X
        Ox = (.Width - MAX_X * L) \ 2
        Oy = .Height - MAX_Y * L
        D = L \ 5
    End With
    
    ' prochaine pièce
    frmMain.pctProchain(1).Width = 4 * L + 10
    frmMain.pctProchain(1).Height = 4 * L + 10
    frmMain.pctProchain(2).Width = 4 * L + 10
    frmMain.pctProchain(2).Height = 4 * L + 10
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : DessineBloc
' Desc : Dessine le bloc (X,Y) dans frmMain.pctFond(Index)
' -----------------------------------------------------------------------------
Public Sub DessineBloc(ByVal X As Single, ByVal Y As Single, _
    ByVal Couleur As ECouleur, ByVal Fond As PictureBox, _
    Optional ByVal ProchainePiece As Boolean = False)
Dim P(3) As POINTAPI
Dim cX As Integer, cY As Integer

    ' origine du dessin
    If ProchainePiece Then
        cX = 4: cY = 4
    Else
        cX = Ox: cY = Oy
    End If
    
    ' couleurs
    ChargeCouleurs Couleur
    
    ' position des sommets
    P(1).X = (X - 1) * L + cX: P(1).Y = (Y - 1) * L + cY
    P(2).X = P(1).X: P(2).Y = P(1).Y + L - 1
    P(3).X = P(1).X + L - 1: P(3).Y = P(2).Y
    
    ' partie arrière (foncée)
    Fond.Line (P(1).X, P(1).Y)-(P(3).X, P(3).Y), C3, BF
    ' partie éclairée
    Fond.FillColor = C1: Fond.ForeColor = C1
    Polygon Fond.hDC, P(1), 3
    ' partie principale
    Fond.Line (P(1).X + D, P(1).Y + D)-(P(3).X - D, P(3).Y - D), C2, BF
    
End Sub

' -----------------------------------------------------------------------------
' Nom  : ChargeCouleurs
' Rem  : C1 = couleur la plus claire; C3 = la plus foncée (sauf pour ORANGE)
' -----------------------------------------------------------------------------
Private Sub ChargeCouleurs(ByVal Couleur As ECouleur)

    Select Case Couleur
        Case ROUGE: C1 = RGB(255, 50, 50): C2 = &HD0&: C3 = &H80&
        Case BLEU: C1 = &HFF8080: C2 = &HC00000: C3 = &H700000
        Case JAUNE: C1 = &H50FFFF: C2 = RGB(230, 230, 0): C3 = &H8080&
        Case GRIS: C1 = &HE0E0E0: C2 = &HA0A0A0: C3 = &H707070
        Case VERT: C1 = &HC000&: C2 = &H8000&: C3 = &H3000&
        Case NOIR: C1 = vbBlack: C2 = vbRed: C3 = vbYellow
    End Select
    
End Sub

