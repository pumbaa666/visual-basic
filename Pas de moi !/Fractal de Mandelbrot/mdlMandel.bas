Attribute VB_Name = "mdlMandel"
Option Explicit

'LES FRACTALS (pris de www.bibmath.net)
'http://www.bibmath.net/dico/index.php3?action=affiche&quoi=./m/mandelbrot.html

 '"Soit c un nombre complexe et fc la fonction définie par fc(z)=z²+c.
 'Si z0=0, on peut définir une suite récurrente par Z(indice n+1)=fc(Zn).
 'L'ensemble de Mandelbrot est l'ensemble des c tels que la suite zn soit bornée.
 'C'est un des ensembles fractales les plus célèbres, découvert par Benoit Mandelbrot,
 'sur lequel on sait encore (relativement) peu de choses."

'Donc en gros, on prend un point du plan complexe d'affixe C.
'On crée une suite récurrente définie par :
'{ Z(n+1) = Z(n)² + C
'{ Z(0) = 0
'Et on regarde si la suite est bornée, c'est à dire si sa limite quand n tend ver l'infini est un réel.
'Comment sait-on quand elle est bornée ?

 '"On calcule la suite (zn) jusqu'à ce que son module dépasse 2
 '(on démontre que c'est équivalent au fait qu'elle diverge).
 'En fonction du premier terme N pour lequel cela se produit, on affecte une couleur au point d'affixe C"

'Alors voilà la fonction qui est à la base de tout... elle calcule ce nombre N.

Public Function MandelbrotNum(ByVal C As clsCpx, ByVal Iter As Integer)
Dim Z As clsCpx 'Le complexe Z
Dim n As Integer    'Pour la boucle
Set Z = New clsCpx  'On définit Z comme étant un complexe   (initialement nul)
For n = 0 To Iter - 1   'jusqu'au nombre d'itérations maximal,
    Z.Square    'on élève Z au carré,
    Z.Add C     'puis on lui ajoute C
    If Z.ModuleCrr > 4 Then 'si son module dépasse 2 (càd si son module au carré dépasse 4 - le carré du module est plus rapide à calculer)
        Exit For    'quitter la boucle
    End If
Next n
MandelbrotNum = n   'et voilà !
End Function

'Je sais, un module entier pour 20 lignes de code c'est pas top, mais ça sert :)
