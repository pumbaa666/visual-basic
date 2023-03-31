VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public angle As Double
Public pos As New pointapi
Public vie As Double
Public Sub init()
    vie = 101
End Sub
Public Sub placer(p1 As pointapi, p2 As pointapi)
    angle = anguler(p1, p2)
    pos.X = p1.X
    pos.Y = p1.Y
    vie = 0
End Sub
Public Sub analyser()
    If vivant Then vie = vie + 0.2
End Sub
Public Sub dessiner()
    analyser
    liner pos, pos2, 2000 - vie
End Sub
Public Function pos2() As pointapi
    Dim p As New pointapi
    p.X = pos.X + Cos(angle * pi / 180)
    p.Y = pos.Y + Sin(angle * pi / 180)
    Set pos2 = p
End Function
Public Function vivant() As Boolean
    vivant = vie < 100
End Function
