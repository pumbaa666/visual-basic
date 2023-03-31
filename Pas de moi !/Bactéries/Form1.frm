VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   9675
   FillColor       =   &H000080FF&
   FillStyle       =   5  'Downward Diagonal
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   645
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'_______________________________
'
' Rejoignez le projet F2X
' al_iksir@hotmail.com
' http://www.actualiteo.com
'_______________________________

'on crée un tableau avec au maximim 20 bactéries
Private bact(20) As New bacterie
'et un tableau avec au maximum 20 aliments
Private alim(20) As New aliment

Private Sub Form_Load()
    Me.WindowState = 2
    Form1.FillStyle = 0
    For i = 0 To UBound(bact)
        'on initialise les bactéries sur l'écran au hasard
        bact(i).init Me.ScaleWidth * Rnd, Me.ScaleHeight * Rnd, Rnd * 20, Rnd * 360
    Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'vous pouvez placer un aliment sur l'ecran en cliquant
    placeraliment X, Y
End Sub
Public Sub placeraliment(X, Y, Optional ray As Long = 20)
    Dim d As Integer: d = 0
    Dim fini As Boolean: fini = False
    While Not fini
        'on parcourt le tableau des aliments
        'on cherche le premier aliment mort (mangé)
        If d <= UBound(alim) And Not alim(d).vivant Then
            'quand on l'a trouvé on arrete le parcours du tableau
            fini = True
            'on place l'aliment à l'endroit indiqué
            alim(d).init CLng(X), CLng(Y), Rnd * ray + 2
            'on lui donne la vie
            alim(d).vivant = True
        ElseIf d = UBound(alim) Then
            'si on est arrivé a la fin du tableau
            'on arrete le parcours
            'il n'y a pas assez d'aliments disponibles
             fini = True
        End If
        d = d + 1
    Wend
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 1 Then placeraliment X, Y
End Sub

Private Sub Timer1_Timer()
    'on efface l'ecran
    Me.Cls
    For i = 0 To UBound(alim)
        'pour chaque aliment vivant de l'ecran ...
        If alim(i).vivant Then
            For k = 0 To UBound(alim)
                'on regarde les autres aliments vivants de l'ecran
                If alim(k).vivant Then
                    'et on essaie de les faire fusionner
                    alim(k).fusionner alim(i)
                End If
            Next k
            'puis on dessine l'aliment
            alim(i).dessiner
            For j = 0 To UBound(bact)
                'pour chaque bactérie de l'ecran
                'on essaie de la nourrir avec l'aliment courant
                alim(i).nourrir bact(j)
            Next j
        End If
    Next i
    For i = 0 To UBound(bact)
        'de temps en temps (1 fois sur 3)
        'on offre a la bactérie une poussée comprise entre 0 et 2
        If Rnd * 3 < 1 Then bact(i).avancer Rnd * 2
        'on fait tourner la bactérie de 5° à droite ou a gauche au hasard
        'en effet, le nombre [rnd * 10 - 5] est compris entre -5 et 5
        bact(i).tourner Rnd * 10 - 5
        'puis on dessine la bactérie
        bact(i).dessiner
    Next i
End Sub
