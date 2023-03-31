VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctC 
      Height          =   135
      Left            =   2880
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   2640
      Width           =   135
   End
   Begin VB.PictureBox pctB 
      Height          =   135
      Left            =   1560
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   1440
      Width           =   135
   End
   Begin VB.PictureBox pctA 
      Height          =   135
      Left            =   840
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   3240
      Width           =   135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dragging As Integer
Dim Xm As Integer
Dim Ym As Integer

Private Sub Form_Click()
If Dragging > -1 Then
Select Case Dragging
Case 0
pctA.Top = Ym
pctA.Left = Xm
pctA.BackColor = RGB(0, 0, 255)
Case 1
pctB.Top = Ym
pctB.Left = Xm
pctB.BackColor = RGB(0, 0, 255)
Case 2
pctC.Top = Ym
pctC.Left = Xm
pctC.BackColor = RGB(0, 0, 255)
End Select
Dragging = -1
Bezier
End If
End Sub

Private Sub Form_activate()
Dragging = -1
pctA.BackColor = RGB(0, 0, 255)
pctB.BackColor = RGB(0, 0, 255)
pctC.BackColor = RGB(0, 0, 255)
Bezier
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xm = X
Ym = Y
End Sub

Private Sub pctA_Click()
If Dragging < 0 Then
Dragging = 0
pctA.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub pctB_Click()
If Dragging < 0 Then
Dragging = 1
pctB.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub pctC_Click()
If Dragging < 0 Then
Dragging = 2
pctC.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub Bezier()
'La fonction principale ! Tout tient sur une seule ligne.
'Les courbes de bézier sont des courbes utilisées en imagerie de synthèse et dans le tracé de police de caractère,
'C'est plus ou moins des droite qu'on a fait "tordre" avec l'ajout d'un troisème point... C'est du moins l'impression
'que j'en ai...
'Voici comment on définit une de ces courbes :
'On a trois points A, B et C dans un repère d'origine O.
'[Dans ce programme, O se situe au coin haut-gauche du formulaire]
'La courbe est décrite par le point M, tel que :
'->              ->                     ->        ->
'OM = (1 - t)² * OA + 2 * t * (1 - t) * OB + t² * OC
'Pour t variant de 0 à 1. Voilà !

'D'après Math 1reS Programme 2001, collection Hyperbole, Nathan.

Dim t As Double
Me.Refresh
For t = 0 To 1 Step 0.001
Me.PSet ((1 - t) * (1 - t) * pctA.Left + 2 * t * (1 - t) * pctB.Left + t * t * pctC.Left, (1 - t) * (1 - t) * pctA.Top + 2 * t * (1 - t) * pctB.Top + t * t * pctC.Top), 0
Next t
End Sub
