Attribute VB_Name = "Module1"
'// Note : certaines fonctions, comme la fonction :
'//          CreateRoundRectRgn
'// et :
'//           CreatePolygonRgn
'// ne sont pas utilis� dans cette exemple.
'// Cela dit, il vous suffira de les tester de la m�me
'// mani�re, ce n'est pas tr�s difficile...

'// Type n�cessaire aux r�gions polygonales :
Type POINTAPI
    X As Long
    Y As Long
End Type

'// Les fonctions permettant de cr�er des r�gions invisibles
'// sur la form.

'// R�gions rectangulaires :
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'// R�gions circulaires ou en ellipses :
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'// R�gions polygonales :
Declare Function CreatePolygonRgn Lib "gdi32" (IpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'// R�gions rectangulaires avec bords arondis :
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'// Fonction permettant d'associ� plusieurs r�gions :
Declare Function CombineRgn Lib "gdi32" (ByVal hdestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'// Fonction permettant d'appliquer les r�gions sur une form :
Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'// Fonction permettant de lib�rer la m�moire :
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'// Constantes d'op�rateurs logiques :

'// ET logique :
Public Const RGN_AND = 1
'// OU logique :
Public Const RGN_OR = 2
'// OU exclusif :
Public Const RGN_XOR = 3
'// Soustraction logique :
Public Const RGN_DIFF = 4
'// Si vous n'avez pas trop train� la grole dans vos ann�es
'// de lyc�e, vous devriez pouvoir vous rappellez de ce
'// qu'est un intervalle, une diff�rence, une union,
'// ou une intersection... Et ben c'est la m�me chose !
'// Et faite moi pas croire que vous n'avez jamais appris
'// �a, je suis e seconde et je le sais d�j� ! :p

'// Tables de v�rit� de ces op�rateurs :
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
