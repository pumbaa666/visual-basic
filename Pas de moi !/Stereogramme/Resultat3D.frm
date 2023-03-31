VERSION 5.00
Begin VB.Form Resultat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stéréo3D"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   2385
   ClientWidth     =   6375
   Icon            =   "Resultat3D.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   6375
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   120
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Resultat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim px As Single, py As Single

Private Sub Form_Activate()
    FormActive = Me.Caption
    MDI3D.CD3D.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, unloadmode As Integer)
    FormActive = ""
    MDI3D.CD3D.Enabled = False
End Sub

Private Sub Form_Load()
    FormActive = Me.Caption
    'Image1.Stretch = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
x As Single, y As Single)
    FormActive = Me.Caption
    If Button = 2 Then   ' Vérifiez si le bouton droit de la souris
                         ' a été actionné.
        MDI3D.PopupMenu MDI3D.menucache2
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Tag = "*": px = x: py = y
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Deplace Image1, x, y
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Tag = ""
    FormActive = Me.Caption
    If Button = 2 Then   ' Vérifiez si le bouton droit de la souris
                        ' a été actionné.
        MDI3D.PopupMenu MDI3D.menucache2
    End If

End Sub

Sub Deplace(F As Object, x As Single, y As Single)
    If F.Tag = "*" Then
        F.Tag = "": F.Move F.Left + (x - px), F.Top + (y - py)
        F.ZOrder 0: F.Tag = "*"
    End If
End Sub

