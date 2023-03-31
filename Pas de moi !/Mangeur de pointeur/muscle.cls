VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "muscle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pos As New pointapi
Public angle As Double
Public longueur As Double
Public courant As Double
Public vitesse As Double
Public amplitude As Double
Public poussee As Double
Public Sub init(Optional X As Double = 200, Optional Y As Double = 200, Optional longu As Double = 10, Optional angl As Double = 45, Optional crn As Double = 0, Optional ampl As Double = 10, Optional vitesse2 As Double = 0.5)
    pos.X = X
    pos.Y = Y
    longueur = longu
    angle = angl
    courant = crn
    vitesse = vitesse2
    amplitude = ampl
    poussee = 0
End Sub
Public Sub analyser()
    positionner
    avancer vitesse
    angle = moduler(angle)
    tourner Rnd * 10 - 5
    courant = courant + 0.1
    'If courant > 12 * pi Then courant = 0
End Sub
Public Sub avancer(Optional nb As Double = 1)
    poussee = poussee + nb
    If poussee > 5 Then poussee = 5
End Sub
Public Sub tourner(Optional nb As Double = 1)
    angle = angle + nb
End Sub
Public Sub dessiner(Optional color = vbBlack)
    liner pos, pos2, color
End Sub
Public Function pos2() As pointapi
    Dim p As New pointapi
    p.X = pos.X + longueur * Cos(angle2 * pi / 180)
    p.Y = pos.Y + longueur * Sin(angle2 * pi / 180)
    Set pos2 = p
End Function
Public Function angle2() As Double
    angle2 = moduler(angle + amplitude * Cos(courant / (vitesse * 0.5)))
End Function
Public Sub positionner()
    poussee = poussee * frottement
    pos.X = pos.X - vitesse * Cos(angle2 * pi / 180)
    pos.Y = pos.Y - vitesse * Sin(angle2 * pi / 180)
    '-------------------------------------
    replacer pos
End Sub

