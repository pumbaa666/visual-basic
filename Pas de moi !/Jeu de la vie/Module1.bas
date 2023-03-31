Attribute VB_Name = "Module1"
Dim vie(100, 100)
Dim vie_(100, 100)
Dim vie__(10000, 2)
Global bmin, bmax, brev



Sub lavie()
'*** Variable des limites
bmin = 2
bmax = 3
brev = 3
'*** Numero de la picturebox en cours
scr = 0
'
Form1.gfx(0).ForeColor = vbRed
Form1.gfx(1).ForeColor = vbRed
Randomize
'** Initialise les cellules en vie
'regenvie
'
Boucle_de_la_vie:
'******************************* On calcul la vie
For X = 2 To 99
For Y = 2 To 99
'** On recupere les valeurs autour de la
'** cellule en cours
a1 = vie(X - 1, Y + 1)
a2 = vie(X, Y + 1)
a3 = vie(X + 1, Y + 1)
a4 = vie(X - 1, Y)
a5 = vie(X, Y)
a6 = vie(X + 1, Y)
a7 = vie(X - 1, Y - 1)
a8 = vie(X, Y - 1)
a9 = vie(X + 1, Y - 1)
'** La somme de toutes les cases
som = a1 + a2 + a3 + a4 + a6 + a7 + a8 + a9
'** 1er test de la vie : La Naissance
'** Une cellule mort devient vivante si elle a
'** exactement trois cellules voisines vivantes
'**
'** Et elle morte ?
If a5 = 0 Then
    '** a t elle le droit de revivre
    If som = brev Then
    a5 = 1 '** Oui
    End If
Else
    '** 2eme test de la vie : La Survie
    '** Une cellule reste en vie tant qu'elle a deux
    '** ou trois voisines vivantes
    '** Pas morte Donc elle est en vie
    If som <> bmin And som <> bmax Then
    a5 = 0 '** Elle meurt
    Else
    '** 3eme test de la vie : La Mort
    '** Dans les autres cas la cellule meurt par
    '** etouffement (+de 3 voisines vivantes )
    '** ou de solitude ( - de 2 voisines )
    '**
       If som > bmax Or som < bmin Then a5 = 0
    '**
    End If
End If
'
'** On rempli le tableau virtuel de la vie
'** avec la nouvelle valeur
'** pas le meme par ca fausserait les resultats
vie_(X, Y) = a5
'** Cellule de la colonne suivante
Next
'** Cellule de la Ligne suivante
Next
'** On remplace le vrai tableau de la vie
'** par le virtuel
For X = 1 To 100
For Y = 1 To 100
If Form1.Option1.Value = Checked Then vie(X, Y) = vie_(X, Y)
Next
Next
'******************************* On affiche la vie
Form1.gfx(scr).Cls '** Efface le screen gfx
'** On affiche les cellules
c = 0
For X = 1 To 100
For Y = 1 To 100
If vie(X, Y) = 1 Then Form1.gfx(scr).PSet (X * 4, Y * 4)
Next
Next
'
'*** On affiche la picture box tempo
Form1.gfx(scr).Visible = True
'*** On inverse la valeur de scr 0 ou 1
scr = -scr + 1
'*** On casse la futur picture box ou l on va dessiner
Form1.gfx(scr).Visible = False
'
'***************************************************
DoEvents
GoTo Boucle_de_la_vie



End Sub

Sub addcell(X, Y)
'*** Crée un nid de vie
Debug.Print X, Y
'*** Jamais + de 100 et pas - de 0
X = 1 + Abs(Int(X / 4) Mod 99)
Y = 1 + Abs(Int(Y / 4) Mod 99)
vie(X - 1, Y + 1) = 1
vie(X, Y + 1) = 1
vie(X + 1, Y + 1) = 1
vie(X - 1, Y) = 1
vie(X, Y) = 1
vie(X + 1, Y) = 1
vie(X - 1, Y - 1) = 1
vie(X, Y - 1) = 1
vie(X + 1, Y - 1) = 1
End Sub

Sub regenvie()
'*** Creation de 1500 points aléatoirement
For t = 1 To 1500
X = 25 + Int(Rnd * 50)
Y = 25 + Int(Rnd * 50)
vie(X, Y) = 1
Next
End Sub


Sub vide()
'*** Vide la vrai vie
For X = 1 To 100
For Y = 1 To 100
vie(X, Y) = 0
Next
Next
End Sub
