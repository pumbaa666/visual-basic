Attribute VB_Name = "mdlMandel"
Option Explicit

'LES FRACTALS (pris de www.bibmath.net)
'http://www.bibmath.net/dico/index.php3?action=affiche&quoi=./m/mandelbrot.html

 '"Soit c un nombre complexe et fc la fonction d�finie par fc(z)=z�+c.
 'Si z0=0, on peut d�finir une suite r�currente par Z(indice n+1)=fc(Zn).
 'L'ensemble de Mandelbrot est l'ensemble des c tels que la suite zn soit born�e.
 'C'est un des ensembles fractales les plus c�l�bres, d�couvert par Benoit Mandelbrot,
 'sur lequel on sait encore (relativement) peu de choses."

'Donc en gros, on prend un point du plan complexe d'affixe C.
'On cr�e une suite r�currente d�finie par :
'{ Z(n+1) = Z(n)� + C
'{ Z(0) = 0
'Et on regarde si la suite est born�e, c'est � dire si sa limite quand n tend ver l'infini est un r�el.
'Comment sait-on quand elle est born�e ?

 '"On calcule la suite (zn) jusqu'� ce que son module d�passe 2
 '(on d�montre que c'est �quivalent au fait qu'elle diverge).
 'En fonction du premier terme N pour lequel cela se produit, on affecte une couleur au point d'affixe C"

'Alors voil� la fonction qui est � la base de tout... elle calcule ce nombre N.

Public Function MandelbrotNum(ByVal C As clsCpx, ByVal Iter As Integer)
Dim Z As clsCpx 'Le complexe Z
Dim n As Integer    'Pour la boucle
Set Z = New clsCpx  'On d�finit Z comme �tant un complexe   (initialement nul)
For n = 0 To Iter - 1   'jusqu'au nombre d'it�rations maximal,
    Z.Square    'on �l�ve Z au carr�,
    Z.Add C     'puis on lui ajoute C
    If Z.ModuleCrr > 4 Then 'si son module d�passe 2 (c�d si son module au carr� d�passe 4 - le carr� du module est plus rapide � calculer)
        Exit For    'quitter la boucle
    End If
Next n
MandelbrotNum = n   'et voil� !
End Function

'Je sais, un module entier pour 20 lignes de code c'est pas top, mais �a sert :)
