Attribute VB_Name = "Module1"
'_______________________________
'
' Rejoignez le projet F2X
' al_iksir@hotmail.com
' http://www.actualiteo.com
'_______________________________

Public Const pi As Long = 3.14159265
Public Function pow(nb As Long) As Double
    pow = CDbl(nb * nb)
End Function
Public Function moduler(nb As Long) As Long
    If nb < 0 Then
        nb = 360 + nb
    ElseIf nb > 360 Then
        nb = nb - 360
    End If
    moduler = nb
End Function
