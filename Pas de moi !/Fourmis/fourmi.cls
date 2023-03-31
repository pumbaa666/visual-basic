VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fourmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public manger As New pointapi
Public posa As New pointapi
Public pos As New pointapi
Public angle As Double
Public longueur As Double
Public male As Boolean
Public charge As Boolean
Public Sub init(X As Double, Y As Double)
    posa.X = X
    posa.Y = Y
    pos.X = X
    pos.Y = Y
    angle = moduler(Rnd * 360)
    longueur = 1 + Rnd * 3
    male = IIf(Rnd * 100 < 50, True, False)
    manger.X = Rnd * 250
    manger.Y = Rnd * 250
    charge = False
End Sub
Public Sub charger()
    charge = True
End Sub
Public Sub decharger()
    charge = False
End Sub
Public Sub manger_est(p As pointapi)
    manger.X = p.X
    manger.Y = p.Y
End Sub
Public Sub maison_est(p As pointapi)
    maison.X = p.X
    maison.Y = p.Y
End Sub
Public Sub analyser()
    If pos.X < 0 Then pos.X = Form1.ScaleWidth
    If pos.X > Form1.ScaleWidth Then pos.X = 0
    If pos.Y < 0 Then pos.Y = Form1.ScaleHeight
    If pos.Y > Form1.ScaleHeight Then pos.Y = 0
    '----------------------------
    'manger.X = manger.X + Rnd * 2 - 1: manger.Y = manger.Y + Rnd * 2 - 1
    If charge And Rnd * 100 < longueur * 500 Then Form1.feromoner pos, manger
End Sub
Public Sub tracer()
    analyser
    If dist(posa, pos) < longueur * 2 Then liner posa, pos, IIf(male, vbRed, vbWhite)
    'liner manger, pos, vbBlue
    If charge Then cercler posa, 1, vbWhite
End Sub

Public Sub avancer()
    posa.X = pos.X
    posa.Y = pos.Y
    pos.X = pos.X + longueur * Cos(angle * pi / 180)
    pos.Y = pos.Y + longueur * Sin(angle * pi / 180)
    '---------------------
    If charge Then
        angle = anguler(Form1.maison.pos, pos)
    Else
        angle = angle + Rnd * 10 - 5
    End If
End Sub
Public Function femelle() As Boolean
    femelle = Not male
End Function
