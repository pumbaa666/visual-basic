VERSION 5.00
Begin VB.Form FrmMenu 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casse 130"
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6525
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FrmMenu.frx":4E7E
   MousePointer    =   1  'Arrow
   ScaleHeight     =   4650
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   3810
      Left            =   3360
      ScaleHeight     =   3750
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   3060
      Begin VB.Image pionex 
         Height          =   750
         Index           =   19
         Left            =   750
         Picture         =   "FrmMenu.frx":5B48
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Image pionex 
         Height          =   1500
         Index           =   18
         Left            =   1500
         Picture         =   "FrmMenu.frx":6024
         Top             =   2250
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   1500
         Index           =   17
         Left            =   0
         Picture         =   "FrmMenu.frx":6551
         Top             =   2250
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   1500
         Index           =   16
         Left            =   750
         Picture         =   "FrmMenu.frx":6A7E
         Top             =   2250
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   1500
         Index           =   15
         Left            =   2250
         Picture         =   "FrmMenu.frx":6FAB
         Top             =   2250
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   750
         Index           =   14
         Left            =   2250
         Picture         =   "FrmMenu.frx":74D8
         Top             =   750
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   750
         Index           =   13
         Left            =   0
         Picture         =   "FrmMenu.frx":7963
         Top             =   750
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   750
         Index           =   12
         Left            =   0
         Picture         =   "FrmMenu.frx":7DEE
         Top             =   1500
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   750
         Index           =   11
         Left            =   2250
         Picture         =   "FrmMenu.frx":8279
         Top             =   1500
         Width           =   750
      End
      Begin VB.Image pionex 
         Height          =   1500
         Index           =   10
         Left            =   750
         Picture         =   "FrmMenu.frx":8704
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   3810
      Left            =   120
      ScaleHeight     =   3750
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   3060
      Begin VB.Image Pion 
         Height          =   1500
         Index           =   1
         Left            =   750
         Picture         =   "FrmMenu.frx":8E42
         Top             =   2250
         Width           =   1500
      End
      Begin VB.Image Pion 
         Height          =   750
         Index           =   7
         Left            =   1500
         Picture         =   "FrmMenu.frx":9580
         Top             =   750
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   750
         Index           =   8
         Left            =   750
         Picture         =   "FrmMenu.frx":9A0B
         Top             =   1500
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   750
         Index           =   6
         Left            =   750
         Picture         =   "FrmMenu.frx":9E96
         Top             =   750
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   750
         Index           =   9
         Left            =   1500
         Picture         =   "FrmMenu.frx":A321
         Top             =   1500
         Width           =   748
      End
      Begin VB.Image Pion 
         Height          =   1500
         Index           =   5
         Left            =   2250
         Picture         =   "FrmMenu.frx":A7AC
         Top             =   2250
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   1500
         Index           =   2
         Left            =   0
         Picture         =   "FrmMenu.frx":ACD9
         Top             =   750
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   1500
         Index           =   3
         Left            =   0
         Picture         =   "FrmMenu.frx":B206
         Top             =   2250
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   1500
         Index           =   4
         Left            =   2250
         Picture         =   "FrmMenu.frx":B733
         Top             =   750
         Width           =   750
      End
      Begin VB.Image Pion 
         Height          =   750
         Index           =   0
         Left            =   750
         Picture         =   "FrmMenu.frx":BC60
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Votre objectif :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label LblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu MnuFile 
      Caption         =   "Fichier"
      Begin VB.Menu MnuNew 
         Caption         =   "Nouveau"
      End
      Begin VB.Menu menu3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRep 
         Caption         =   "Voire l'objectif"
         Checked         =   -1  'True
      End
      Begin VB.Menu Menu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "A propos"
      End
      Begin VB.Menu MnuQuitt 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngX As Long
Public lngY As Long
Dim car(64) As Boolean
Dim mouseX, mouseY As Single
Dim Deplace As Integer  'Empéche les déplacement diagonale, 1 pour horizontale, 2 verticale
'                           (uniquemment nécéssaire pour les petit carré)

    'On initialise tous...
Private Sub Form_Load()
  Dim I As Integer
  
   'repére les carrés libre ou non
  For I = 0 To 64
    car(I) = True
  Next I
  car(11) = False
  car(14) = False
End Sub

Private Sub MnuAbout_Click()
  Dim temp As Integer
  temp = MsgBox("Casse téte réalisé par Passaero, tous droits réservé. Vous ne pouvez distribuer ce programme sans l'accord écris de son auteur." & vbCrLf & "Plus d'information : http://www.passaero.tk", vbInformation + vbOKOnly, "Casse130")
End Sub

    'Pour tous recomencer
Private Sub MnuNew_Click()
  Dim I As Integer
  
   'Le compteur est remis à 0
  Timer1.Enabled = True
  LblTime = "00:00:00"
   'les car pour les déplacement sont actualiser
  For I = 0 To 64
    car(I) = True
  Next I
  car(11) = False
  car(14) = False
   'et bien sur les pion sont placé
  Pion(0).Left = 750
  Pion(0).Top = 0
  Pion(1).Left = 750
  Pion(1).Top = 2250
  Pion(2).Left = 0
  Pion(2).Top = 750
  Pion(3).Left = 0
  Pion(3).Top = 2250
  Pion(4).Left = 2250
  Pion(4).Top = 750
  Pion(5).Left = 2250
  Pion(5).Top = 2250
  Pion(6).Left = 750
  Pion(6).Top = 750
  Pion(7).Left = 1500
  Pion(7).Top = 750
  Pion(8).Left = 750
  Pion(8).Top = 1500
  Pion(9).Left = 1500
  Pion(9).Top = 1500
  
End Sub

    'On quitte tous...
Private Sub MnuQuitt_Click()
  End
End Sub

    'On affiche l'objectif ou non
Private Sub MnuRep_Click()
  If MnuRep.Checked = False Then
    FrmMenu.Width = 6615
    MnuRep.Checked = True
   Else
    FrmMenu.Width = 3360
    MnuRep.Checked = False
  End If
End Sub

'La partie de code suivante gére les déplacement des pions....

    'Paramétre quelque variable
Private Sub Pion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  mouseX = X  'pour savoir par ou la sourie se déplace
  mouseY = Y  'méme chose mais sur l'axe Y
  lngX = Pion(Index).Left   'Va permettre de repérer la position du pion
  lngY = Pion(Index).Top    'pareille
  Deplace = 0  'ne sére que pour les petits carré cf Pion_MouseMove
End Sub

Private Sub Pion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim place1 As Integer 'Définie la position du carré en haut a gauche
  Dim temp As Integer 'Pour les message d'erreur
  
  If Not Button = 1 Then Exit Sub 'on verifie que la sourie est cliquée
   'Commençons par récupérer la position du pion
   'Sur l'axe X
  Select Case lngX
   Case 0
     place1 = 1
   Case 750
     place1 = 2
   Case 1500
     place1 = 3
   Case 2250
     place1 = 4
   Case Else
     temp = MsgBox("erreur lngX " & lngX, vbCritical + vbOKOnly, "Casse130")
     Call Pion_MouseUp(Index, 1, 0, 0, 0) 'Pour placer le pion dans une zone normalle
     Exit Sub
  End Select
   'Sur l'axe Y
  Select Case lngY
   Case 0
     place1 = place1 + 10
   Case 750
     place1 = place1 + 20
   Case 1500
     place1 = place1 + 30
   Case 2250
     place1 = place1 + 40
   Case 3000
     place1 = place1 + 50
   Case Else
     temp = MsgBox("erreur lngY " & lngY, vbCritical + vbOKOnly, "Casse130")
     Call Pion_MouseUp(Index, 1, 0, 0, 0) 'Pour placer le pion dans une zone normalle
     Exit Sub
  End Select
  
  
   'on touve le pion selectionné
  Select Case Index
   
   
   Case 0
   '8888888888888888888888888888888888888888888888888888888888888888888888888888888
   ''''''''''''''''''''''''''''''Barre Horizontale'''''''''''''''''''''''''''''''''
      'par ou se fait le déplacement horizontaux
     If mouseX - X < 0 Then 'on va à droite
        If car(place1 + 2) = False Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
        End If
      Else 'on va à gauche
        If car(place1 - 1) = False Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
        End If
     End If
      'par ou se fait le déplacement verticaux
     If mouseY - Y < 0 Then 'on va en bas
        If car(place1 + 10) = False And car(place1 + 11) = False Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
        End If
      Else 'on va en haut
        If car(place1 - 10) = False And car(place1 - 9) = False Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
        End If
     End If
       ''''''''''''''''''''''''''Verification'''''''''''''''''''''''''''''''''''''
       'on vérifie que le pion n'est pas au dela des limites
       'En ca de déplécement rapide la position n'est pas détectée
          Select Case place1 'Gére les déplacements verticaux
           Case 11 To 14
              'on determine ses déplacement max et on corrige
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 750 Then
               Pion(Index).Top = 750
             End If
             'si la case suivante est libre c'est pas grave...
           Case 21 To 24
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 1500 Then
               Pion(Index).Top = 1500
             End If
           Case 31 To 34
             If Pion(Index).Top < 750 Then
               Pion(Index).Top = 750
             End If
             If Pion(Index).Top > 2250 Then
               Pion(Index).Top = 2250
             End If
           Case 41 To 44
             If Pion(Index).Top < 1500 Then
               Pion(Index).Top = 1500
             End If
             If Pion(Index).Top > 3000 Then
               Pion(Index).Top = 3000
             End If
           Case 51 To 54
             If Pion(Index).Top < 2250 Then
               Pion(Index).Top = 2250
             End If
             If Pion(Index).Top > 3000 Then
               Pion(Index).Top = 3000
             End If
           Case Else
             temp = MsgBox("Erreur Verification position X", vbCritical + vbOKOnly, "Casse130")
          End Select
          Select Case place1  'puis on regarde sur l'axe Y et on fait pareille
            Case 11, 21, 31, 41, 51
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 750 Then
                Pion(Index).Left = 750
              End If
            Case 12, 22, 32, 42, 52
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 1500 Then
                Pion(Index).Left = 1500
              End If
            Case 13, 23, 33, 43, 53
              If Pion(Index).Left < 750 Then
                Pion(Index).Left = 750
              End If
              If Pion(Index).Left > 1500 Then
                Pion(Index).Left = 1500
              End If
          End Select
      '''''''''''''''''''''''''''''''Position'''''''''''''''''''''''''''''''''''''
      'on vérifie que l'on est pas dans une position connue et on change les cars
      'On met les deux carré précédament occupé sur false car il sont peut étre libre
     car(place1) = False  'ces lignes sont nécéssaire pour Pion_mouseup
     car(place1 + 1) = False
      'Sur l'axe X
     Select Case Pion(Index).Left
      Case 0
        place1 = 1
      Case 750
        place1 = 2
      Case 1500
        place1 = 3
      Case Else 'le carré n'est pas dans une ce ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
      'Sur l'axe Y
     Select Case Pion(Index).Top
      Case 0
        place1 = place1 + 10
      Case 750
        place1 = place1 + 20
      Case 1500
        place1 = place1 + 30
      Case 2250
        place1 = place1 + 40
      Case 3000
        place1 = place1 + 50
      Case Else 'le carré n'est pas dans une de ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
       'on admet ici que le carré est dans une position connue
     Call Pion_MouseDown(Index, 1, 0, X, Y)
     car(place1) = True
     car(place1 + 1) = True
     
     
   Case 1
   '8888888888888888888888888888888888888888888888888888888888888888888888888888888
   '''''''''''''''''''''''''''''''''LE GROS CARREE'''''''''''''''''''''''''''''''''
      'par ou se fait le déplacement horizontaux
     If mouseX - X < 0 Then 'on va à droite
        If car(place1 + 2) = False And car(place1 + 12) = False Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
        End If
      Else 'on va à gauche
        If car(place1 - 1) = False And car(place1 + 9) = False Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
        End If
     End If
      'par ou se fait le déplacement verticaux
     If mouseY - Y < 0 Then 'on va en bas
        If car(place1 + 20) = False And car(place1 + 21) = False Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
        End If
      Else 'on va en haut
        If car(place1 - 10) = False And car(place1 - 9) = False Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
        End If
     End If
       ''''''''''''''''''''''''''Verification'''''''''''''''''''''''''''''''''''''
       'on vérifie que le pion n'est pas au dela des limites
       'En ca de déplécement rapide la position n'est pas détectée
          Select Case place1 ' on va voire ou est le pion sur l'axe Y
           Case 11 To 14  'on determine ses déplacement max et on corrige
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 750 Then
               Pion(Index).Top = 750
             End If
             'si la case suivante est libre c'est pas grave...
           Case 21 To 24
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 1500 Then
               Pion(Index).Top = 1500
             End If
           Case 31 To 34
             If Pion(Index).Top < 750 Then
               Pion(Index).Top = 750
             End If
             If Pion(Index).Top > 2250 Then
               Pion(Index).Top = 2250
             End If
           Case 41 To 44
             If Pion(Index).Top > 2250 Then
               Pion(Index).Top = 2250
             End If
             If Pion(Index).Top < 1500 Then
               Pion(Index).Top = 1500
             End If
          End Select
          Select Case place1  'puis on regarde sur l'axe Y et on fait pareille
            Case 11, 21, 31, 41, 51
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 750 Then
                Pion(Index).Left = 750
              End If
            Case 12, 22, 32, 42, 52
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 1500 Then
                Pion(Index).Left = 1500
              End If
            Case 13, 23, 33, 43, 53
              If Pion(Index).Left < 750 Then
                Pion(Index).Left = 750
              End If
              If Pion(Index).Left > 2250 Then
                Pion(Index).Left = 2250
              End If
            Case 14, 24, 34, 44, 54
              If Pion(Index).Left < 1500 Then
                Pion(Index).Left = 1500
              End If
              If Pion(Index).Left > 2250 Then
                Pion(Index).Left = 2250
              End If
          End Select
      'Fin de l'ajoutation
      '''''''''''''''''''''''''''''''Position'''''''''''''''''''''''''''''''''''''
      'on vérifie que l'on est pas dans une position connue et on change les cars
      'On met les deux carré précédament occupé sur false car il sont peut étre libre
     car(place1) = False  'c'est ligne de code son nécéssaire pour mouseup
     car(place1 + 1) = False
     car(place1 + 11) = False
     car(place1 + 10) = False
      'Sur l'axe X
     Select Case Pion(Index).Left
      Case 0
        place1 = 1
      Case 750
        place1 = 2
      Case 1500
        place1 = 3
      Case Else 'le carré n'est pas dans une ce ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
      'Sur l'axe Y
     Select Case Pion(Index).Top
      Case 0
        place1 = place1 + 10
      Case 750
        place1 = place1 + 20
      Case 1500
        place1 = place1 + 30
      Case 2250
        place1 = place1 + 40
      Case Else 'le carré n'est pas dans une de ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
       'on admet donc ici que le carré est dans une position connue
     Call Pion_MouseDown(Index, 1, 0, X, Y)
     car(place1) = True
     car(place1 + 1) = True
     car(place1 + 11) = True
     car(place1 + 10) = True
     
     
   Case 2 To 5
   '8888888888888888888888888888888888888888888888888888888888888888888888888888888
   ''''''''''''''''''''''''''''''''BARRE VERTICALE'''''''''''''''''''''''''''''''''
      'Par ou se fait le déplacement
     If mouseX - X < 0 Then 'on va à droite
        If car(place1 + 1) = False And car(place1 + 11) = False Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
        End If
      Else 'on va à gauche
        If car(place1 - 1) = False And car(place1 + 9) = False Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
        End If
     End If
      'par ou se fait le déplacement verticaux
     If mouseY - Y < 0 Then 'on va en bas
        If car(place1 + 20) = False Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
        End If
      Else 'on va en haut
        If car(place1 - 10) = False Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
        End If
     End If
       ''''''''''''''''''''''''''Verification'''''''''''''''''''''''''''''''''''''
       'on vérifie que le pion n'est pas au dela des limites
       'En ca de déplécement rapide la position n'est pas détectée
          Select Case place1 ' on va voire ou est le pion sur l'axe Y
           Case 11 To 14  'on determine ses déplacement max et on corrige
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 750 Then
               Pion(Index).Top = 750
             End If
             'si la case suivante est libre c'est pas grave...
           Case 21 To 24
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 1500 Then
               Pion(Index).Top = 1500
             End If
           Case 31 To 34
             If Pion(Index).Top < 750 Then
               Pion(Index).Top = 750
             End If
             If Pion(Index).Top > 2250 Then
               Pion(Index).Top = 2250
             End If
           Case 41 To 44
             If Pion(Index).Top < 1500 Then
               Pion(Index).Top = 1500
             End If
             If Pion(Index).Top > 3000 Then
               Pion(Index).Top = 3000
             End If
           Case 51 To 54
             If Pion(Index).Top < 2250 Then
               Pion(Index).Top = 2250
             End If
             If Pion(Index).Top > 3000 Then
               Pion(Index).Top = 3000
             End If
          End Select
          Select Case place1  'puis on regarde sur l'axe Y et on fait pareille
            Case 11, 21, 31, 41, 51
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 750 Then
                Pion(Index).Left = 750
              End If
            Case 12, 22, 32, 42, 52
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 1500 Then
                Pion(Index).Left = 1500
              End If
            Case 13, 23, 33, 43, 53
              If Pion(Index).Left < 750 Then
                Pion(Index).Left = 750
              End If
              If Pion(Index).Left > 2250 Then
                Pion(Index).Left = 2250
              End If
            Case 14, 24, 34, 44, 54
              If Pion(Index).Left < 1500 Then
                Pion(Index).Left = 1500
              End If
              If Pion(Index).Left > 2250 Then
                Pion(Index).Left = 2250
              End If
          End Select
      'Fin de l'ajoutation
      '''''''''''''''''''''''''''''''Position'''''''''''''''''''''''''''''''''''''
      'on vérifie que l'on est pas dans une position connue et on change les cars
      'On met les deux carré précédament occupé sur false car il sont peut étre libre
     car(place1) = False
     car(place1 + 10) = False
      'Sur l'axe X
     Select Case Pion(Index).Left
      Case 0
        place1 = 1
      Case 750
        place1 = 2
      Case 1500
        place1 = 3
      Case 2250
        place1 = 4
      Case Else 'le carré n'est pas dans une ce ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
      'Sur l'axe Y
     Select Case Pion(Index).Top
      Case 0
        place1 = place1 + 10
      Case 750
        place1 = place1 + 20
      Case 1500
        place1 = place1 + 30
      Case 2250
        place1 = place1 + 40
      Case Else 'le carré n'est pas dans une de ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
       'on admet donc ici que le carré est dans une position particuliére
     Call Pion_MouseDown(Index, 1, 0, X, Y)
     car(place1) = True
     car(place1 + 10) = True
     
     
   Case 6 To 9
   '8888888888888888888888888888888888888888888888888888888888888888888888888888888
   '''''''''''''''''''''''''''''''LES PETIT CARREE'''''''''''''''''''''''''''''''''
      'par ou se fait le déplacement horizontaux
     If mouseX - X < 0 Then 'on va à droite
        If car(place1 + 1) = False And Deplace <= 1 Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
          Deplace = 1
        End If
      Else 'on va à gauche
        If car(place1 - 1) = False And Deplace <= 1 Then 'le carré est libre
           'on se déplace
          Pion(Index).Left = Pion(Index).Left - (mouseX - X)
          Deplace = 1
        End If
     End If
      'par ou se fait le déplacement verticaux
     If mouseY - Y < 0 Then 'on va en bas
        If car(place1 + 10) = False And Not Deplace = 1 Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
          Deplace = 2
        End If
      Else 'on va en haut
        If car(place1 - 10) = False And Not Deplace = 1 Then 'le carré est libre
          Pion(Index).Top = Pion(Index).Top - (mouseY - Y)
          Deplace = 2
        End If
     End If
       ''''''''''''''''''''''''''Verification'''''''''''''''''''''''''''''''''''''
       'on vérifie que le pion n'est pas au dela des limites
       'En ca de déplécement rapide la position n'est pas détectée
          Select Case place1 ' on va voire ou est le pion sur l'axe Y
           Case 11 To 14  'on determine ses déplacement max et on corrige
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 750 Then
               Pion(Index).Top = 750
             End If
             'si la case suivante est libre c'est pas grave...
           Case 21 To 24
             If Pion(Index).Top < 0 Then
               Pion(Index).Top = 0
             End If
             If Pion(Index).Top > 1500 Then
               Pion(Index).Top = 1500
             End If
           Case 31 To 34
             If Pion(Index).Top < 750 Then
               Pion(Index).Top = 750
             End If
             If Pion(Index).Top > 2250 Then
               Pion(Index).Top = 2250
             End If
           Case 41 To 44
             If Pion(Index).Top < 1500 Then
               Pion(Index).Top = 1500
             End If
             If Pion(Index).Top > 3000 Then
               Pion(Index).Top = 3000
             End If
           Case 51 To 54
             If Pion(Index).Top < 2250 Then
               Pion(Index).Top = 2250
             End If
             If Pion(Index).Top > 3000 Then
               Pion(Index).Top = 3000
             End If
           Case Else
             temp = MsgBox("erreur place1", vbCritical + vbOKOnly, "Casse130")
          End Select
          Select Case place1  'puis on regarde sur l'axe Y et on fait pareille
            Case 11, 21, 31, 41, 51
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 750 Then
                Pion(Index).Left = 750
              End If
            Case 12, 22, 32, 42, 52
              If Pion(Index).Left < 0 Then
                Pion(Index).Left = 0
              End If
              If Pion(Index).Left > 1500 Then
                Pion(Index).Left = 1500
              End If
            Case 13, 23, 33, 43, 53
              If Pion(Index).Left < 750 Then
                Pion(Index).Left = 750
              End If
              If Pion(Index).Left > 2250 Then
                Pion(Index).Left = 2250
              End If
            Case 14, 24, 34, 44, 54
              If Pion(Index).Left < 1500 Then
                Pion(Index).Left = 1500
              End If
              If Pion(Index).Left > 2250 Then
                Pion(Index).Left = 2250
              End If
          End Select
      'Fin de l'ajoutation
      '''''''''''''''''''''''''''''''Position'''''''''''''''''''''''''''''''''''''
      'on vérifie que l'on est pas dans une position connue et on change les cars
      'On met les deux carré précédament occupé sur false car il sont peut étre libre
     car(place1) = False
      'Sur l'axe X
     Select Case Pion(Index).Left
      Case 0
        place1 = 1
      Case 750
        place1 = 2
      Case 1500
        place1 = 3
      Case 2250
        place1 = 4
      Case Else 'le carré n'est pas dans une ce ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
      'Sur l'axe Y
     Select Case Pion(Index).Top
      Case 0
        place1 = place1 + 10
      Case 750
        place1 = place1 + 20
      Case 1500
        place1 = place1 + 30
      Case 2250
        place1 = place1 + 40
      Case 3000
        place1 = place1 + 50
      Case Else 'le carré n'est pas dans une de ces places
        Exit Sub ''''''''''''''''''''''''''''''''''/!!!!!\ ERREUR POSSIBLE
     End Select
       'on admet donc ici que le carré est dans une position particuliére
     Call Pion_MouseDown(Index, 1, 0, X, Y)
     car(place1) = True
     
  End Select
End Sub

Private Sub Pion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim place1 As Integer
  Dim I As Single
  Dim temp As Integer
  
   'On définie la position du pion et on le place bien
  Select Case Pion(Index).Top  'sur l'axe Y
    Case 0 To 375
      Pion(Index).Top = 0
    Case 376 To 1125
      Pion(Index).Top = 750
    Case 1126 To 1875
      Pion(Index).Top = 1500
    Case 1876 To 2625
      Pion(Index).Top = 2250
    Case 2626 To 3375
      Pion(Index).Top = 3000
  End Select
  Select Case Pion(Index).Left  'sur l'axeX
    Case 0 To 375
      Pion(Index).Left = 0
    Case 376 To 1125
      Pion(Index).Left = 750
    Case 1126 To 1875
      Pion(Index).Left = 1500
    Case 1876 To 2625
      Pion(Index).Left = 2250
  End Select
   'on cherche la position du pion
   'sur l'axe X
  Select Case Pion(Index).Left
   Case 0
     place1 = 1
   Case 750
     place1 = 2
   Case 1500
     place1 = 3
   Case 2250
     place1 = 4
   Case Else
     temp = MsgBox("erreur lngX mouseup " & lngX, vbCritical + vbOKOnly, "Casse130")
     Exit Sub
  End Select
   'Sur l'axe Y
  Select Case Pion(Index).Top
   Case 0
     place1 = place1 + 10
   Case 750
     place1 = place1 + 20
   Case 1500
     place1 = place1 + 30
   Case 2250
     place1 = place1 + 40
   Case 3000
     place1 = place1 + 50
   Case Else
     temp = MsgBox("erreur lngY mouseup" & lngY, vbCritical + vbOKOnly, "Casse130")
     Exit Sub
  End Select
  Select Case Index
    Case 0 'barre horizontale
      car(place1) = True
      car(place1 + 1) = True
    Case 1 'gros carré
      car(place1) = True
      car(place1 + 1) = True
      car(place1 + 10) = True
      car(place1 + 11) = True
    Case 2 To 5  'barre verticale
      car(place1) = True
      car(place1 + 10) = True
    Case 6 To 9  'petit carré
      car(place1) = True
    Case Else 'on sais jamais
      MsgBox ("Erreur Mouseup Index " & Index)
      Exit Sub
   End Select
   
   'On regarde si on a pas gagné...
   'On vérifie si chaque pion on bien les coordonnnés prévue dans ce cas
   'Pour la barre horizontale
  If Pion(0).Left <> 750 Or Pion(0).Top <> 1500 Then Exit Sub
   'Pour le gros carré
  If Pion(1).Left <> 750 Or Pion(1).Top <> 0 Then Exit Sub
   'Pour les barres verticales
  For I = 2 To 5
    If Pion(I).Left <> 0 And Pion(I).Left <> 750 And Pion(I).Left <> 1500 And Pion(I).Left <> 2250 Then Exit Sub
    If Pion(I).Top <> 2250 Then Exit Sub
  Next I
   'Pour les petit carré
  For I = 6 To 9
    If Pion(I).Left <> 0 And Pion(I).Left <> 2250 Then Exit Sub
    If Pion(I).Top <> 750 And Pion(I).Top <> 1500 Then Exit Sub
  Next I
   'Si il arrive la c'est que l'on a gagné
  Timer1.Enabled = False
  temp = MsgBox("Tien cela me rappelle quelque chose, vous n'auriez pas gagné ???? " & vbCrLf & "Vous avez mis : " & LblTime, vbApplicationModal + vbOKOnly, "Casse130")
End Sub

    'Mesure du temps
Private Sub Timer1_Timer()
  Dim intH As Integer
  Dim intMin As Integer
  Dim intS As Integer
  Dim strConv As String
  
   'On récupére le temps
  strConv = LblTime.Caption
  intH = CInt(Left(strConv, 2))
  intMin = CInt(Right(Left(strConv, 5), 2))
  intS = CInt(Right(strConv, 2))
  
   'On incrémente les secondes puis on convertit en H min s
  intS = intS + 1
  If intS = 60 Then
    intMin = intMin + 1
    intS = 0
    If intMin = 60 Then
      intH = intH + 1
      intMin = 0
    End If
  End If
  
  If intH < 10 Then
    strConv = "0" & intH & ":"
   Else
    strConv = intH & ":"
  End If
  If intMin < 10 Then
    strConv = strConv & "0" & intMin & ":"
   Else
    strConv = strConv & intMin & ":"
  End If
  If intS < 10 Then
    strConv = strConv & "0" & intS
   Else
    strConv = strConv & intS
  End If
  LblTime.Caption = strConv
End Sub
