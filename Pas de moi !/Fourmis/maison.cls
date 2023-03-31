VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "maison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pos As New pointapi
Public rayon As Double
Public Sub init(Optional X As Double = 100, Optional Y As Double = 100, Optional m_rayon As Double = 5)
    rayon = m_rayon
    pos.X = X
    pos.Y = Y
End Sub
Public Sub dessiner()
    cercler pos, CInt(rayon), vbYellow
End Sub

