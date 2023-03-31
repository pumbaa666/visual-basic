Attribute VB_Name = "Module1"
'// Note : certaines fonctions, comme la fonction :
'//          CreateRoundRectRgn
'// et :
'//           CreatePolygonRgn
'// ne sont pas utilisé dans cette exemple.
'// Cela dit, il vous suffira de les tester de la même
'// manière, ce n'est pas très difficile...

'// Type nécessaire aux régions polygonales :
Type POINTAPI
    X As Long
    Y As Long
End Type

'// Les fonctions permettant de créer des régions invisibles
'// sur la form.

'// Régions rectangulaires :
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'// Régions circulaires ou en ellipses :
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'// Régions polygonales :
Declare Function CreatePolygonRgn Lib "gdi32" (IpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'// Régions rectangulaires avec bords arondis :
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'// Fonction permettant d'associé plusieurs régions :
Declare Function CombineRgn Lib "gdi32" (ByVal hdestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'// Fonction permettant d'appliquer les régions sur une form :
Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'// Fonction permettant de libérer la mémoire :
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'// Constantes d'opérateurs logiques :

'// ET logique :
Public Const RGN_AND = 1
'// OU logique :
Public Const RGN_OR = 2
'// OU exclusif :
Public Const RGN_XOR = 3
'// Soustraction logique :
Public Const RGN_DIFF = 4
'// Si vous n'avez pas trop trainé la grole dans vos années
'// de lycée, vous devriez pouvoir vous rappellez de ce
'// qu'est un intervalle, une différence, une union,
'// ou une intersection... Et ben c'est la même chose !
'// Et faite moi pas croire que vous n'avez jamais appris
'// ça, je suis e seconde et je le sais déjà ! :p

'// Tables de vérité de ces opérateurs :
'RGN_AND :
'       0        1
'0      0        0
'1      0        1

'RGN_OR :
'       0        1
'0      0        1
'1      1        1

'RGN_XOR :
'       0        1
'0      0        1
'1      1        0

'RGN_DIFF :
'       0        1
'0      0        0
'1      1        0
