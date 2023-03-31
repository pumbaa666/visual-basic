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
Public pos As New pointapi
Public rayon As Double
Public angle As Double
Public pris As Boolean
Public Sub init(Optional X As Double = 10, Optional Y As Double = 10, Optional m_rayon As Double = 5)
    rayon = m_rayon
    pos.X = X
    pos.Y = Y
    angle = Rnd * 360
    pris = False
End Sub
Public Sub analyser()
    If rayon <= 1.5 Then
        Me.init Form1.ScaleWidth * Rnd, Form1.ScaleHeight * Rnd
    End If
    avancer
    If pos.X < 0 Then pos.X = Form1.ScaleWidth
    If pos.X > Form1.ScaleWidth Then pos.X = 0
    If pos.Y < 0 Then pos.Y = Form1.ScaleHeight
    If pos.Y > Form1.ScaleHeight Then pos.Y = 0
End Sub
Public Sub dessiner()
    analyser
    cercler pos, CInt(rayon)
End Sub
Public Sub avancer()
    'pos.X = pos.X + 0.2 * Cos(angle * pi / 180)
    'pos.Y = pos.Y + 0.2 * Sin(angle * pi / 180)
    '-------------
    'angle = angle + Rnd * 10 - 5
End Sub
