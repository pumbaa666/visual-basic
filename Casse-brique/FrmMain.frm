VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casse-briques"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FormAireJeu 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9495
      Begin VB.Shape ShpBrique 
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   675
      End
      Begin VB.Shape ShpRaquette 
         Height          =   255
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   7680
         Width           =   975
      End
      Begin VB.Shape ShpBalle 
         Height          =   255
         Left            =   5880
         Shape           =   3  'Circle
         Top             =   7440
         Width           =   255
      End
   End
   Begin VB.Timer ClkBalle 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   8640
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   8640
      Width           =   1695
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NbY = 7
Const NbX = 8
Dim tBrique(NbX, NbX) As Integer
Dim vDeltaX As Integer
Dim vDeltaY As Integer
Dim vNbBrique As Integer

Private Sub ClkBalle_Timer()
Dim vExit As Boolean

    ShpBalle.Left = ShpBalle.Left + vDeltaX
    ShpBalle.Top = ShpBalle.Top - vDeltaY

    ' Si la balle touche le haut de l'écran
    If ShpBalle.Top <= FormAireJeu.Top Then
        vDeltaY = vDeltaY * -1
    End If

    ' Si la balle touche les bords de l'écran
    If ShpBalle.Left <= 0 Or ShpBalle.Left + ShpBalle.Width / 2 >= FormAireJeu.Width - FormAireJeu.Left Then
        vDeltaX = vDeltaX * -1
    End If

    ' Si la balle arrive au niveau de la raquette
    If (ShpBalle.Top + ShpBalle.Height) >= ShpRaquette.Top Then
        If ShpBalle.Left + ShpBalle.Width >= ShpRaquette.Left And ShpBalle.Left <= ShpRaquette.Left + ShpRaquette.Width Then
            vDeltaY = vDeltaY * -1

            ' Si la balle est à gauche de la raquette
            If ShpBalle.Left + ShpBalle.Width / 2 < ShpRaquette.Left + ShpRaquette.Width / 2 Then
                vDeltaX = ((ShpRaquette.Left + ShpRaquette.Width / 2) - (ShpBalle.Left + ShpBalle.Width / 2)) / 10
                If vDeltaX > 0 Then vDeltaX = vDeltaX * -1
            Else ' Si elle est à droite
                vDeltaX = ((ShpRaquette.Left + ShpRaquette.Width / 2) - (ShpBalle.Left + ShpBalle.Width / 2)) / 10
                If vDeltaX < 0 Then vDeltaX = vDeltaX * -1
            End If
        Else
            'MsgBox "Perdu"
            ClkBalle.Enabled = False
        End If
    End If

    ' Si la balle touche une brique
    For vCount = 0 To vNbBrique
        If (ShpBalle.Left <= ShpBrique(vCount).Left + ShpBrique(vCount).Width) And (ShpBalle.Left > ShpBrique(vCount).Left) And (ShpBalle.Top + ShpBalle.Height / 2 <= ShpBrique(vCount).Top + ShpBrique(vCount).Height) And (ShpBalle.Top + ShpBalle.Height / 2) >= ShpBrique(vCount).Top Then ' Si la balle touche une brique par la droite
            If tBrique(Int(vCount / NbY) - 1, NbY - vCount Mod NbY - 1) <> 0 Then
                vDeltaX = vDeltaX * -1
                vExit = True
                ShpBrique(vCount).Visible = False
                tBrique(Int(vCount / NbY) - 1, NbY - vCount Mod NbY - 1) = 0
            End If
'               Exit For
        End If

        If (ShpBalle.Left + ShpBalle.Width >= ShpBrique(vCount).Left) And (ShpBalle.Left + ShpBalle.Width < ShpBrique(vCount).Left) And (ShpBalle.Top + ShpBalle.Height / 2 <= ShpBrique(vCount).Top + ShpBrique(vCount).Height) And (ShpBalle.Top + ShpBalle.Height / 2) >= ShpBrique(vCount).Top Then ' Si la balle touche une brique par la gauche
            If tBrique(Int(vCount / NbY) - 1, NbY - vCount Mod NbY - 1) <> 0 Then
                vDeltaX = vDeltaX * -1
                vExit = True
                ShpBrique(vCount).Visible = False
                tBrique(Int(vCount / NbY) - 1, NbY - vCount Mod NbY - 1) = 0
'               Exit For
            End If
        End If

        If (ShpBalle.Top <= ShpBrique(vCount).Top + ShpBrique(vCount).Height) And (ShpBalle.Top > ShpBrique(vCount).Top) And (ShpBalle.Left <= ShpBrique(vCount).Left + ShpBrique(vCount).Width) And (ShpBalle.Left + ShpBalle.Width >= ShpBrique(vCount).Left) Then  ' Si la balle touche une brique par le bas
            If tBrique(Int(vCount / NbY) - 1, NbY - vCount Mod NbY - 1) <> 0 Then
                If vExit = False Then
                    vDeltaY = vDeltaY * -1
                    vExit = True
                    ShpBrique(vCount).Visible = False
                    tBrique(Int(vCount / NbY) - 1, NbY - vCount Mod NbY - 1) = 0
                End If
'               Exit For
            End If
        End If
        If vExit = True Then
            Exit For
        End If
    Next
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdQuitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeplacementSouris X + CmdQuitter.Left
End Sub

Private Sub FormAireJeu_Click()
    If ClkBalle.Enabled = False Then
        ClkBalle.Enabled = True
        vDeltaX = Int(Rnd * 20) - 10
        vDeltaY = 40
    End If
End Sub

Private Sub Form_Load()
Dim vCount As Integer
Dim vCount2 As Integer

    Randomize
    For vCount = 0 To NbX
        For vCount2 = 1 To NbY
            Load ShpBrique(vCount * NbY + vCount2)
            ShpBrique(vCount * NbY + vCount2).Left = vCount * ShpBrique(vCount * NbY + vCount2).Width
            ShpBrique(vCount * NbY + vCount2).Top = vCount2 * ShpBrique(vCount * NbY + vCount2).Height + FormAireJeu.Top
            ShpBrique(vCount * NbY + vCount2).Visible = True
            tBrique(vCount, vCount2) = 1
            vNbBrique = vNbBrique + 1
        Next
    Next
    ShpBrique(0).Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeplacementSouris X
End Sub

Function DeplacementSouris(ByVal vX As Integer)
    ' Permet de ne pas dépasser l'aire de jeu avec la raquette
    If vX < FormAireJeu.Width - ShpRaquette.Width / 2 And vX > ShpRaquette.Width / 2 Then
        ShpRaquette.Left = vX - (ShpRaquette.Width / 2)
        If ClkBalle.Enabled = False Then
            ShpBalle.Left = vX - (ShpBalle.Width / 2)
        End If
    End If
End Function

Private Sub FormAireJeu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DeplacementSouris X + FormAireJeu.Left
End Sub
