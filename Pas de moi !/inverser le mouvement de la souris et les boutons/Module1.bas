Attribute VB_Name = "Module1"
'D�claration d'un nouveau type de variable contenant 2 �l�ments
Public Type Orthonorm�
    X As Long
    Y As Long
End Type


'Fonction permettant de conna�tre la position X et Y du curseur de la souris
Declare Function GetCursorPos Lib "user32" (lpPoint As Orthonorm�) As Long

'Fonction permettant de d�finir des coordonn�es pour le curseur de la souris
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Declare Function SwapMouseButton& Lib "user32" (ByVal bSwap As Long)

