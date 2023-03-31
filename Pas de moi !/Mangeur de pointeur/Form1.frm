VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   664
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   907
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public d As New pointapi
Private m(0 To 50) As New muscle


Private Sub Form_Load()
    d.X = 0
    d.Y = 0
    For i = 0 To UBound(m)
        m(i).init Me.ScaleWidth * Rnd, Me.ScaleHeight * Rnd, 10 + Rnd * 5, Rnd * 360, Rnd, 10 + Rnd * 10, 2 + Rnd
    Next i
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     d.X = X
     d.Y = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For i = 0 To UBound(m)
        Set m(i) = Nothing
    Next i
End Sub

Private Sub Timer1_Timer()
    Me.Cls
    Dim pris As Boolean: pris = False
    If replacer(d, 10) Then SetCursorPos d.X, d.Y
    For i = 0 To UBound(m)
        m(i).analyser
        ang = anguler(m(i).pos, d)
        If moduler(m(i).angle - ang) > 180 Then
            m(i).tourner 2
        Else
            m(i).tourner -2
        End If
        If Not pris Then
            If (dist(m(i).pos, d) < m(i).longueur * 2 Or dist(m(i).pos2, d) < m(i).longueur * 2) Then
                'cercler m(i).pos, m(i).longueur, vbRed
                d.X = m(i).pos.X + Cos(m(i).angle * pi / 180)
                d.Y = m(i).pos.Y + Sin(m(i).angle * pi / 180)
                SetCursorPos d.X, d.Y
                pris = True
            End If
        End If
        m(i).dessiner
    Next i
End Sub

