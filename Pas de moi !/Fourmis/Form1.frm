VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Les fourmis"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   13935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Les fourmis"
   ScaleHeight     =   654
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   929
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private f(0 To 50) As New fourmi
Private fero(0 To 300) As New fero
Private alim(0 To 5) As New aliment
Public maison As New maison
Private Sub Form_Load()
    Randomize Timer
    Me.BackColor = vbBlack
    For i = 0 To UBound(f)
        f(i).init Me.ScaleWidth * Rnd, Me.ScaleHeight * Rnd
    Next i
    For i = 0 To UBound(fero)
        fero(i).init
    Next i
    For i = 0 To UBound(alim)
        alim(i).init Rnd * Me.ScaleWidth, Rnd * Me.ScaleHeight
    Next i
    maison.init
End Sub
Public Sub feromoner(pos1 As pointapi, pos2 As pointapi)
    Dim fini As Boolean: fini = False
    i = 0
    While i < UBound(fero) And Not fini
        If fero(i).vivant Then
            If dist(fero(i).pos, pos1) < 5 Then
                fero(i).vie = 0
                fini = True
            End If
        Else
            fero(i).placer pos1, pos2
            fini = True
        End If
        i = i + 1
    Wend
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim p As New pointapi
    p.X = X
    p.Y = Y
    If Button = 1 Then
        For i = 0 To UBound(alim)
            If dist(p, alim(i).pos) < alim(i).rayon * 3 Then
                alim(i).pris = True
            End If
        Next i
        'alim.init CDbl(X), CDbl(Y)
    ElseIf Button = 2 Then
        maison.init CDbl(X), CDbl(Y)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        For i = 0 To UBound(alim)
            If alim(i).pris Then
                alim(i).pos.X = X
                alim(i).pos.Y = Y
            End If
        Next i
        'alim.init CDbl(X), CDbl(Y)
    ElseIf Button = 2 Then
        maison.init CDbl(X), CDbl(Y)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        For i = 0 To UBound(alim)
            alim(i).pris = False
        Next i
End Sub

Private Sub Timer1_Timer()
    Me.Cls
    For i = 0 To UBound(alim)
        alim(i).dessiner
    Next i
    maison.dessiner
    For i = 0 To UBound(fero)
        fero(i).dessiner
    Next i
    For i = 0 To UBound(f)
        For a = 0 To UBound(alim)
            diista = dist(f(i).pos, alim(a).pos)
            If diista < (f(i).longueur + alim(a).rayon) * 2 Then
                liner alim(a).pos, f(i).pos, vbWhite
                f(i).manger_est alim(a).pos
                f(i).charger
                alim(a).rayon = alim(a).rayon - 0.01
            End If
        Next a
        If Not f(i).charge Then
            For k = 0 To UBound(fero)
                If fero(k).vivant Then
                    If Rnd * 100 < 100 - fero(k).vie Then
                        diistf = dist(fero(k).pos, f(i).pos)
                        If diistf < (100 - fero(k).vie) / 10 Then
                            f(i).angle = moduler(fero(k).angle + 180)
                            liner fero(k).pos, f(i).pos, RGB(100, 0, 0)
                        End If
                    End If
                End If
            Next k
        End If
        diistm = dist(f(i).pos, maison.pos)
        If diistm < (f(i).longueur + maison.rayon) Then
            liner maison.pos, f(i).pos, vbYellow
            f(i).decharger
        End If
        f(i).avancer
        f(i).tracer
    Next i
End Sub

Public Sub commentaire()
        For j = 0 To UBound(f)
            diist = dist(f(i).pos, f(j).pos)
            If diist < (f(i).longueur + f(j).longueur) * 2 Then
                If f(i).male And f(j).femelle Then
                    liner f(i).pos, f(j).pos, vbRed
                    f(i).manger_est f(j).manger
                End If
            End If
        Next j
End Sub
