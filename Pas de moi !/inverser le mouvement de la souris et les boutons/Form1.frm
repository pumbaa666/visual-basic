VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inverseur de souris"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Tout remettre"
      Height          =   855
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   5895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tout inverser"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Inverser les boutons de la souris"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restaurer les boutons de la souris"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Arréter l'inversion"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Démarrer l'inversion"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pos_Précédente As Orthonormé 'Contient la position du curseur capturée une 1ère fois
Dim Pos_Actuelle As Orthonormé 'Contient la position du curseur capturée une 2ème fois, 1 demi 100e de seconde plus tard
Dim Pos_Inverse As Orthonormé 'Contient la position inverse du curseur par rapport à sa "Pos_Précédente" et sa "Pos_Actuelle"

Private Sub Command3_Click()
    SwapMouseButton False

End Sub

Private Sub Command4_Click()
    SwapMouseButton True
End Sub

Private Sub Command5_Click()
    SwapMouseButton True
    'Capture les coordonnées du curseur
    GetCursorPos Pos_Précédente
    'Copie ses coordonnées dans une 2e variable
    Pos_Inverse = Pos_Précédente
    'Active la minuterie qui contient le code
    Timer1.Enabled = True

End Sub

Private Sub Command6_Click()
    Timer1.Enabled = False
SwapMouseButton False

End Sub

Private Sub Form_Load()
    'Initialise l'objet Timer
    Timer1.Enabled = False
    Timer1.Interval = 5
End Sub

Private Sub Command1_Click()
    'Capture les coordonnées du curseur
    GetCursorPos Pos_Précédente
    'Copie ses coordonnées dans une 2e variable
    Pos_Inverse = Pos_Précédente
    'Active la minuterie qui contient le code
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    'Arréte l'inversion du mouvement du curseur en désactivant l'objet Timer
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    'Enregistre la position actuelle du curseur
    GetCursorPos Pos_Actuelle


    'Teste si le curseur bouge, si c'est le cas il calcule le mouvement inverse et l'applique au curseur
    If Pos_Actuelle.X <> Pos_Précédente.X Or Pos_Actuelle.Y <> Pos_Précédente.Y Then
        'Calcul du mouvement inverse
        Pos_Inverse.X = Pos_Précédente.X + ((Pos_Actuelle.X - Pos_Précédente.X) * (-1))
        Pos_Inverse.Y = Pos_Précédente.Y + ((Pos_Actuelle.Y - Pos_Précédente.Y) * (-1))
        'positionne le curseur
        SetCursorPos Pos_Inverse.X, Pos_Inverse.Y
    End If


    'Teste si le curseur a atteint l'un des 4 bords de l'écran, si c'est le cas
    'le curseur ressort de l'autre côté de l'écran (cela l'empêche de rester coincé)
    '
    'Si touche le bord droit de l'écran le curseur est replacé sur le bord gauche
    If Pos_Inverse.X >= ScaleX(Screen.Width, vbTwips, vbPixels) Then SetCursorPos 1, Pos_Actuelle.Y
    'Pareil pour le bord gauche
    If Pos_Inverse.X <= 1 Then SetCursorPos ScaleX(Screen.Width, vbTwips, vbPixels) - 2, Pos_Actuelle.Y
    'Pareil pour le bas de l'écran
    If Pos_Inverse.Y >= ScaleY(Screen.Height, vbTwips, vbPixels) Then SetCursorPos Pos_Actuelle.X, 1
    'Pareil pour le haut de l'écran
    If Pos_Inverse.Y <= 1 Then SetCursorPos Pos_Actuelle.X, ScaleY(Screen.Height, vbTwips, vbPixels) - 2
    '
    'Les méthodes ScaleX ou ScaleY convertissent une mesure de hauteur ou de largeur en une autre


    'Enregistre la position du curseur pour pouvoir la comparer à la prochaine "position actuelle" et en déduire le mouvement de la souris
    GetCursorPos Pos_Précédente
End Sub
