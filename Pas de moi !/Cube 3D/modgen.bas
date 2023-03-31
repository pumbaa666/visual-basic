Attribute VB_Name = "modgen"
Const Xorig = 2100
Const Yorig = 2100
Const focal = -1500
Const cubcot = 1000

Public mode As Integer
Public angle As Single

Type point3D
    x As Integer
    y As Integer
    z As Integer
End Type

Type point2D
    x As Integer
    y As Integer
End Type

Type cube
    p(7) As point3D
End Type

Public cub As cube
Function projection(arf As point3D) As point2D
    Dim tmp As point2D
    
    tmp.x = Xorig + arf.x * (focal / (focal - arf.z))
    tmp.y = Yorig + arf.y * (focal / (focal - arf.z))
    
    projection = tmp
End Function
Function init_cub()
    cub.p(0).x = cubcot / 2
    cub.p(0).y = cubcot / 2
    cub.p(0).z = cubcot / 2

    cub.p(1).x = cubcot / 2
    cub.p(1).y = -(cubcot / 2)
    cub.p(1).z = cubcot / 2

    cub.p(2).x = -(cubcot / 2)
    cub.p(2).y = -(cubcot / 2)
    cub.p(2).z = cubcot / 2

    cub.p(3).x = -(cubcot / 2)
    cub.p(3).y = cubcot / 2
    cub.p(3).z = cubcot / 2

    cub.p(4).x = cubcot / 2
    cub.p(4).y = cubcot / 2
    cub.p(4).z = -(cubcot / 2)

    cub.p(5).x = cubcot / 2
    cub.p(5).y = -(cubcot / 2)
    cub.p(5).z = -(cubcot / 2)

    cub.p(6).x = -(cubcot / 2)
    cub.p(6).y = -(cubcot / 2)
    cub.p(6).z = -(cubcot / 2)

    cub.p(7).x = -(cubcot / 2)
    cub.p(7).y = cubcot / 2
    cub.p(7).z = -(cubcot / 2)
    
    angle = 0
End Function
Function draw_cub()
    Dim i As Integer
    Dim ptmp(7) As point2D
    
    For i = 0 To 7
        ptmp(i) = projection(cub.p(i))
        main.pic.PSet (ptmp(i).x, ptmp(i).y), RGB(0, 0, 0)
    Next i
    
    Select Case mode
        Case 0:     ' points
            
        Case 1:     ' wireframe
            main.pic.Line (ptmp(0).x, ptmp(0).y)-(ptmp(1).x, ptmp(1).y)
            main.pic.Line (ptmp(1).x, ptmp(1).y)-(ptmp(2).x, ptmp(2).y)
            main.pic.Line (ptmp(2).x, ptmp(2).y)-(ptmp(3).x, ptmp(3).y)
            main.pic.Line (ptmp(3).x, ptmp(3).y)-(ptmp(0).x, ptmp(0).y)
                  
            main.pic.Line (ptmp(4).x, ptmp(4).y)-(ptmp(5).x, ptmp(5).y)
            main.pic.Line (ptmp(5).x, ptmp(5).y)-(ptmp(6).x, ptmp(6).y)
            main.pic.Line (ptmp(6).x, ptmp(6).y)-(ptmp(7).x, ptmp(7).y)
            main.pic.Line (ptmp(7).x, ptmp(7).y)-(ptmp(4).x, ptmp(4).y)
                  
            main.pic.Line (ptmp(0).x, ptmp(0).y)-(ptmp(4).x, ptmp(4).y)
            main.pic.Line (ptmp(1).x, ptmp(1).y)-(ptmp(5).x, ptmp(5).y)
            main.pic.Line (ptmp(2).x, ptmp(2).y)-(ptmp(6).x, ptmp(6).y)
            main.pic.Line (ptmp(3).x, ptmp(3).y)-(ptmp(7).x, ptmp(7).y)
        Case 2:     ' couleur
            face_color ptmp(0).x, ptmp(0).y, ptmp(3).x, ptmp(3).y, ptmp(1).x, ptmp(1).y, ptmp(2).x, ptmp(2).y, RGB(145, 13, 14)
            face_color ptmp(3).x, ptmp(3).y, ptmp(7).x, ptmp(7).y, ptmp(2).x, ptmp(2).y, ptmp(6).x, ptmp(6).y, RGB(13, 150, 14)
            face_color ptmp(7).x, ptmp(7).y, ptmp(4).x, ptmp(4).y, ptmp(6).x, ptmp(6).y, ptmp(5).x, ptmp(5).y, RGB(13, 13, 150)
            face_color ptmp(4).x, ptmp(4).y, ptmp(0).x, ptmp(0).y, ptmp(5).x, ptmp(5).y, ptmp(1).x, ptmp(1).y, RGB(200, 200, 0)
        Case 3:     ' texture
    End Select
          
End Function
Function face_color(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, x3 As Integer, y3 As Integer, x4 As Integer, y4 As Integer, color As Long)
    Dim t As Integer, Ymin As Integer, Ymax As Integer, yT As Integer, xMax As Integer, xMin As Integer
    Dim pt1 As point2D, pt2 As point2D
    
    If x1 < x2 Then
        If y3 < y4 Then
            Ymin = y3
            Ymax = y1
            yT = y4
            pt1.x = x2
            pt2.x = x1
        Else
            Ymin = y4
            Ymax = y2
            yT = y3
            pt1.x = x1
            pt2.x = x2
        End If
        
        For t = 0 To (Ymax - Ymin)
            If y1 < y2 Then
                pt1.y = y1 - t
                pt2.y = y2 - t
            Else
                pt1.y = y2 - t
                pt2.y = y1 - t
            End If
            
            If pt1.y < yT Then
                pt1.y = yT
            End If
            
            main.pic.Line (pt1.x, pt1.y)-(pt2.x, pt2.y), color
            t = t + 1
        Next t
    End If
    
End Function
Function rotat_cub()
    Dim i As Integer
    Dim tmp As point3D
    
    For i = 0 To 7
        'tmpz.x = Cos(angle) * cub.p(i).x - Sin(angle) * cub.p(i).y
        'tmpz.y = Sin(angle) * cub.p(i).x + Cos(angle) * cub.p(i).y
        'tmpz.z = cub.p(i).z
        
        'tmpx.x = cub.p(i).x
        'tmpx.y = Cos(angle) * cub.p(i).y - Sin(angle) * cub.p(i).z
        'tmpx.z = Sin(angle) * cub.p(i).y + Cos(angle) * cub.p(i).z
        
        'tmpy.x = Cos(angle) * cub.p(i).x - Sin(angle) * cub.p(i).z
        'tmpy.y = cub.p(i).y
        'tmpy.z = Sin(angle) * cub.p(i).x + Cos(angle) * cub.p(i).z

        tmp.x = Cos(angle) * cub.p(i).x - Sin(angle) * cub.p(i).z
        tmp.y = cub.p(i).y
        tmp.z = Sin(angle) * cub.p(i).x + Cos(angle) * cub.p(i).z
        
        cub.p(i).x = tmp.x
        cub.p(i).y = tmp.y
        cub.p(i).z = tmp.z
    Next i
    
    If (angle >= 360) Then
        angle = 0
    Else
        angle = CSng((ax + 0.1))
    End If

End Function
