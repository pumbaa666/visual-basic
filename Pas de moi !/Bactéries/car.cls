VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bacterie"
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

'les coordonn�es de la bact�rie
Public X As Long
Public Y As Long
'la longueur de la bact�rie
Public longueur As Long
'la direction du deplacement de la bact�rie
Public angle As Long
'la pouss�e pour permettre d'avancer
Public pouss As Double

Public Sub init(Optional xx As Long = 50, Optional yy As Long = 50, Optional longu As Long = 5, Optional angl As Long = 10)
    X = xx
    Y = yy
    longueur = longu
    angle = angl
    pouss = 0
End Sub

Public Sub dessiner()
    
    'on fait avancer la bact�rie en fonction de la pouss�e
    X = X + pouss * Cos(angle * pi / 180)
    Y = Y + pouss * Sin(angle * pi / 180)
    
    'si la bact�rie sort de l'�cran alors
    'on la fait apparaire a l'oppos�
    If X < 0 Then X = Form1.ScaleWidth
    If X > Form1.ScaleWidth Then X = 0
    If Y < 0 Then Y = Form1.ScaleHeight
    If Y > Form1.ScaleHeight Then Y = 0
    
    'la pouss�e est limit�e par un coefficient de frottement de 0.9
    pouss = pouss * CDbl(0.9)
    
    'on dessine le corps de la bact�rie
    Form1.Line ( _
        X - (longueur / 2) * Cos(angle * pi / 180), _
        Y - (longueur / 2) * Sin(angle * pi / 180) _
    )-( _
        X + (longueur / 2) * Cos(angle * pi / 180), _
        Y + (longueur / 2) * Sin(angle * pi / 180) _
    )
    Form1.FillColor = vbYellow
    Form1.ForeColor = vbYellow
    
    'on dessine le cercle du centre
    Form1.Circle (X, Y), 2
    'puis l'arriere
    Form1.Circle (X - (longueur / 2) * Cos(angle * pi / 180), Y - (longueur / 2) * Sin(angle * pi / 180)), 2
    Form1.FillColor = vbRed
    'puis la tete
    Form1.Circle (X + (longueur / 2) * Cos(angle * pi / 180), Y + (longueur / 2) * Sin(angle * pi / 180)), 2
    '-----------------------
    'de temps en temps, c'est � dire que :
    'on choisit un chiffre entre 0 et 100 et s'il est inf�rieur a 5
    'donc on a 5% de chances que �a se produise
    If Rnd * 100 < 5 Then
        'on place un aliment a l'arriere de la bact�rie
        '(une dejection)
        Form1.placeraliment _
            X - (longueur * 1.5) * Cos(angle * pi / 180), _
            Y - (longueur * 1.5) * Sin(angle * pi / 180), _
            longueur
        'cela diminue la taille de la bact�rie
        Me.longueur = Me.longueur * 0.8
    End If
    'si la taille de la bact�rie est >20
    'alors on consid�re qu'il est mort (trop vieux ou trop ob�se, a vous de voir)
    If longueur > 20 Then
        'donc on replace la bact�rie ailleurs (resurection) avec une taille entre 0 et 5 : (5*rnd)
        Me.init Form1.ScaleWidth * rd, Form1.ScaleHeight * Rnd, 5 * Rnd, 360 * Rnd
    End If
End Sub
Public Sub avancer(nb As Long)
    'pour faire avancer la bact�rie on lui offre une pouss�e
    pouss = pouss + nb
    'cette pouss�e est limit�e � 10 en avant
    If pouss > 10 Then
        pouss = 10
    'et 5 en arriere
    ElseIf pouss < -5 Then
        pouss = -5
    End If
End Sub

Public Sub tourner(nb As Long)
    'on ne peut faire tourner une bact�rie que si sa pouss�e est assez grande
    If Abs(pouss) > 2 Then angle = moduler(angle + nb)
End Sub
