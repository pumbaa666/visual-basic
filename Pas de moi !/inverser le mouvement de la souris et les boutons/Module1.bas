Attribute VB_Name = "Module1"
'Déclaration d'un nouveau type de variable contenant 2 éléments
Public Type Orthonormé
    X As Long
    Y As Long
End Type


'Fonction permettant de connaître la position X et Y du curseur de la souris
Declare Function GetCursorPos Lib "user32" (lpPoint As Orthonormé) As Long

'Fonction permettant de définir des coordonnées pour le curseur de la souris
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Declare Function SwapMouseButton& Lib "user32" (ByVal bSwap As Long)

