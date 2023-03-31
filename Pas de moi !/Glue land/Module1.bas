Attribute VB_Name = "Module1"



Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type


Declare Function IntersectRect Lib "user32" ( _
                 lpDestRect As RECT, _
                 lpSrc1Rect As RECT, _
                 lpSrc2Rect As RECT) As Long


Public vitessetop
Public vitesseleft
Public vitessetop2
Public vitesseleft2

Type asdfghj
left As Variant
top As Variant
End Type
Public bv As asdfghj
Public bv2 As asdfghj


