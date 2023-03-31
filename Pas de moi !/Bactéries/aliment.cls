VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aliment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'_______________________________
'
' Rejoignez le projet F2X
' al_iksir@hotmail.com
' http://www.actualiteo.com
'_______________________________

'les coordonnées de l'aliment
Public X As Double
Public Y As Double
'l'aliment est il deja mangé
Public vivant As Boolean
'le rayon de l'aliment
Public rayon As Double

' Initialisation de l'aliment
Public Sub init(xx As Long, yy As Long, Optional ray As Integer = 20)
    X = xx
    Y = yy
    rayon = ray
    vivant = False
End Sub

Public Sub dessiner()
    'If Me.rayon > 20 Then Me.rayon = 20
    If Me.rayon > 0 Then
        'Form1.FillStyle = 1
        Form1.ForeColor = vbBlue
        Form1.FillColor = vbBlue
        Form1.Circle (X, Y), rayon
        'de temps en temps on fait grandir l'aliment
        'on prend un nombre au hasard entre 0 et le rayon
        'si ce nombre est plus petit que 1 alors on augmrnte la taille de l'aliment
        'donc plus le rayon est grand moins il a de chanche de grandir
        If Rnd * rayon < 1 Then Me.rayon = Me.rayon * 1.045
    Else
        'si le rayon est <= 0 alors l'aliment est mangé (mort)
        Me.vivant = False
    End If
End Sub

'procedure pour savoir si cet aliment est assez proche d'une bacterie (bac) pour la nourrir
Public Sub nourrir(bac As bacterie)
    Dim dist As Long
    With bac
        ' on calcule la distance entre l'aliment et la tete de la bacterie
        ' les coordonnées de l'aliment son X et Y
        ' les coordonnées de la tete de la bactérie sont :
        ' coordX = (.X + (.longueur / 2) * Cos(.angle * pi / 180))
        ' coordY = (.Y + (.longueur / 2) * Sin(.angle * pi / 180))
        dist = Sqr(pow(X - (.X + (.longueur / 2) * Cos(.angle * pi / 180))) + pow(Y - (.Y + (.longueur / 2) * Sin(.angle * pi / 180))))
    End With
    If dist < rayon Then
        'si la distance entre la tete de la bactérie et l'aliment est inférieure au rayon de l'aliment
        'alors on diminue la taille de l'aliment car celui ci est mangé
        Me.rayon = Me.rayon - 2
        'et on augmente la taille de la bactérie qui l'a mangé
        bac.longueur = bac.longueur + 1
    End If
    'si le rayon de l'aliment est <2 alors on considere qu'il est mangé donc mort
    If Me.rayon < 2 Then Me.vivant = False
End Sub

'procedure pour fusionner deux aliments assez proches l'un de l'autre
Public Sub fusionner(al As aliment)
    Dim dist As Double
    If Me.rayon < 1 Then Me.rayon = 1
    With al
        'on calcule la distance entre les 2 aliments
        dist = CDbl(Sqr(pow(X - .X) + pow(Y - .Y)))
        'si la distance est <= a la somme des rayons des aliments
        'donc si les aliments se touchent
        If dist <= .rayon + Me.rayon Then
            'si l'autre aliment est plus gros que l'aliment courant
            If .rayon > Me.rayon Then
                'alors on augmente la taille de l'autre aliment
                .rayon = Sqr(pow(.rayon) + pow(Me.rayon) - pow(Me.rayon - 1))
                'et on diminue la taille de l'aliment courant
                Me.rayon = Me.rayon - 1
            End If
        End If
        'ce qui suit permet de faire s'attirer deux aliments assez proches l'un de l'autre
        'cela permet de donner l'impression que le plus gros gobe le plus petit
        'comme deux bulles qui s'aspirent
        '-> si la distance entre les deux aliments est assez proche
        'et que les aliments sont assez gros (fonction du rayon)
        If dist < (.rayon + Me.rayon) * 2 Then
            'alors on rapproche l'aliment courant de l'autre
            'plus cet aliment est loin et moins il est attiré
            'est inversement
            Me.X = Me.X - 0.01 * (Me.X - .X)
            Me.Y = Me.Y - 0.01 * (Me.Y - .Y)
            'et on rapproche aussi l'autre aliment de l'aliment courant
            .X = .X - 0.01 * (.X - Me.X)
            .Y = .Y - 0.01 * (.Y - Me.Y)
        End If
    End With
End Sub
