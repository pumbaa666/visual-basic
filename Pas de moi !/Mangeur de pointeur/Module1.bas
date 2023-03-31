Attribute VB_Name = "Module1"
Declare Function SetCursorPos Lib "user32" ( _
                 ByVal X As Long, _
                 ByVal Y As Long) As Long
Public Const pi As Double = 3.14159265
Public Const frottement As Double = 0.9
Public Function pow(nb As Double) As Double
    pow = CDbl(nb * nb)
End Function
Public Function moduler(nb As Double) As Double
    If nb < 0 Then
        nb = 360 + nb
    ElseIf nb > 360 Then
        nb = nb - 360
    End If
    moduler = nb
End Function

Public Function pit(p As pointapi, p1 As pointapi, p2 As pointapi, p3 As pointapi) As Boolean
    Dim AB As Double
    Dim BC As Double
    Dim CA As Double
    Dim pit2 As Boolean
    
    'Dim d As Integer:d = 2:Form1.Circle (x, y), d: Form1.Circle (x1, y1), d: Form1.Circle (x2, y2), d: Form1.Circle (x3, y3), d
    AB = ((p.Y - p1.Y) * (p2.X - p1.X)) - ((p.X - p1.X) * (p2.Y - p1.Y))
    BC = ((p.Y - p2.Y) * (p3.X - p2.X)) - ((p.X - p2.X) * (p3.Y - p2.Y))
    If AB * BC <= 0 Then
        pit2 = False
    Else
        pit2 = BC * ((p.Y - p3.Y) * (p1.X - p3.X) - (p.X - p3.X) * (p1.Y - p3.Y)) > 0
    End If
    'If pit2 Then: Form1.Line (x, y)-(x1, y1): Form1.Line (x, y)-(x2, y2): Form1.Line (x, y)-(x3, y3):
    pit = pit2
End Function
Public Function replacer(pos As pointapi, Optional bord As Double = 0) As Boolean
    Dim ok As Boolean: ok = False
    If pos.X < bord Then pos.X = Form1.ScaleWidth - bord: ok = True
    If pos.X > Form1.ScaleWidth - bord Then pos.X = bord: ok = True
    If pos.Y < bord Then pos.Y = Form1.ScaleHeight - bord: ok = True
    If pos.Y > Form1.ScaleHeight - bord Then pos.Y = bord: ok = True
    replacer = ok
End Function
Public Function anguler(p1 As pointapi, p2 As pointapi, Optional diff As Double = 0, Optional coeff As Double = 1) As Long
    Dim alpha As Long
    If (p1.X - p2.X) <> 0 Then
        alpha = Atn((p1.Y - p2.Y) / (p1.X - p2.X)) / pi * 180
    Else
        If p1.Y < p2.Y Then
            alpha = 270
        Else
            alpha = 90
        End If
    End If
    If p1.X < p2.X Then
        If p1.Y < p2.Y Then
            alpha = alpha + 180
        ElseIf p1.Y > p2.Y Then
            alpha = alpha + 180
        Else
            alpha = 180
        End If
    ElseIf p1.X > p2.X Then
        If p1.Y < p2.Y Then
            alpha = alpha + 360
        End If
    Else
        'alpha = 90
    End If
    anguler = moduler(coeff * alpha + diff)
End Function
Public Sub cercler(p As pointapi, taille As Integer, Optional color = vbWhite)
    Form1.Circle (p.X, p.Y), taille, color
End Sub
Public Sub liner(p1 As pointapi, p2 As pointapi, Optional color = vbRed)
    Form1.Line (p1.X, p1.Y)-(p2.X, p2.Y), color
End Sub
Public Function dist(p1 As pointapi, p2 As pointapi) As Double
    dist = Sqr(pow(p1.X - p2.X) + pow(p1.Y - p2.Y))
End Function
Public Function distdroite(pos1 As pointapi, pos2 As pointapi, pos3 As pointapi) As Double
    Dim angl1 As Double
    Dim angl2 As Double
    Dim angl3 As Double
    angl1 = anguler(pos2, pos3)
    angl2 = anguler(pos2, pos1)
    angl3 = Abs(angl1 - angl2)
    distdroite = Abs(dist(pos1, pos2) * Sin(angl3 * pi / 180))
End Function
Public Function timestamp() As Double
    timestamp = Hour(Now) * 10000 + Minute(Now) * 100 + Second(Now)
End Function

