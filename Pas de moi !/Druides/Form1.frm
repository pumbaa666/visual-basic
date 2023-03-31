VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Le Voyageur"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Facultées"
      Height          =   1335
      Left            =   6480
      TabIndex        =   5
      Top             =   7320
      Width           =   3975
      Begin VB.ListBox faculte 
         Height          =   1035
         ItemData        =   "Form1.frx":076A
         Left            =   120
         List            =   "Form1.frx":076C
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Santé"
      Height          =   1335
      Left            =   2040
      TabIndex        =   4
      Top             =   7320
      Width           =   3975
      Begin VB.Label Label7 
         Caption         =   "Nbr de poissons :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nbr de champignons :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label5 
         Caption         =   "Somme :"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nbr de fruits :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Points de vie :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   7920
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6705
      Left            =   240
      Picture         =   "Form1.frx":076E
      ScaleHeight     =   6645
      ScaleWidth      =   12000
      TabIndex        =   2
      Top             =   120
      Width           =   12060
      Begin VB.Image Image1 
         Height          =   210
         Left            =   4440
         Picture         =   "Form1.frx":40DB0
         Top             =   4920
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6705
      Left            =   240
      Picture         =   "Form1.frx":4118E
      ScaleHeight     =   6645
      ScaleWidth      =   12000
      TabIndex        =   1
      Top             =   120
      Width           =   12060
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   12210
      TabIndex        =   3
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   6960
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nage, Vol, Peche, SavoirChampi As Boolean   'Déclaration de variable
Dim OnMarchand, OnDruide, OnPêcheur As Boolean
Public Somme, Life, QtéFruit, QtéChampignon, QtéPoissons As Integer
Dim X, Y
Public PreviousColor As Long

Private Sub Form_Load()
    Life = 1000  'Vie à 100%
    Somme = 1000    'Somme à 1000
    QtéFruit = 0 'Nbr de fruits
    QtéChampignon = 0 'Nbr de fruits
    QtéPoissons = 0 'Nbr de fruits
    Nage = False
    Vol = False
    MsgBox "Pour commencer, cliquer sur la carte", vbInformation, "Le voyageur"
End Sub

Private Sub Picture2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii    'Selon la touche presser on déplace le smiley
        Case 122    'Haut - Z
            Image1.Top = Image1.Top - 10
            Life = Life - 0.1 'Coût en vie pour le déplacement
        Case 115    'Droite - S
            Image1.Left = Image1.Left + 10
            Life = Life - 0.1
        Case 113    'Gauche - Q
            Image1.Left = Image1.Left - 10
            Life = Life - 0.1
        Case 119    'Bas    - W
            Image1.Top = Image1.Top + 10
            Life = Life - 0.1
        Case 32    'Espace - Action
            Randomize   'Lancement du Random
            If Picture1.Point(Image1.Left, Image1.Top) = 16777215 Then    'Si l'on est sur un point Blanc alors...
                Taux = Int((Rnd * 3) + 1)   'Défini le taux de change
                If QtéFruit >= 0 Then   'Si on a des fruits
                    qté = InputBox("Combien de fruits voulez-vous vendre ?" & Chr(13) & "1 Fruit = " & Taux, "Le voyageur", QtéFruit)    'Combien en vend t'on
                    If qté <= QtéFruit Then 'Vérifie que le nombre de fruits ne dépasse pas le stock
                        Somme = Somme + (Val(qté) * Taux)   'Ajout de la somme
                        QtéFruit = QtéFruit - qté   'On enlève le nombre de fruits utilisé
                    Else
                        MsgBox "Vente annulé vous n'avez pas assez de fruits !", vbExclamation, "Le voyageur" 'Pas assez de fruits
                    End If
                Else
                    MsgBox "Vous n'avez rien à vendre", vbExclamation, "Le voyageur" 'Rien à vendre
                End If
            ElseIf Picture1.Point(Image1.Left, Image1.Top) = 5056006 Then    'Si l'on est en pleine mer
                Qt = Int((Rnd * 2) + 0) 'Nombre de poissons ramasser
                MsgBox "Nombre de poissons pêcher : " & Qt, vbInformation, "Le voyageur"
                QtéPoissons = QtéPoissons + Qt    'Ajout des fruits dans le sac
            ElseIf Picture1.Point(Image1.Left, Image1.Top) = 1080840 Then   'Si l'on est en foret
                Qt = Int((Rnd * 5) + 1) 'Nombre de fruits ramasser
                MsgBox "Nombre de fruits ramasser : " & Qt, vbInformation, "Le voyageur"
                QtéFruit = QtéFruit + Qt    'Ajout des fruits dans le sac
            ElseIf Picture1.Point(Image1.Left, Image1.Top) = 1310465 Then   'Si l'on est dans un bosquet
                Qt = Int((Rnd * 3) + 1) 'Nombre de Champignons ramasser
                MsgBox "Nombre de champignons ramasser : " & Qt, vbInformation, "Le voyageur"
                If SavoirChampi = False Then    'Test di l'on connais les champignons
                    If Int((Rnd * (Qt * 2)) + 1) = Qt Then 'Détermine si on tombe sur un mauvais champignon soit 1 chance sur 2 par champignons
                        MsgBox "Vous avez ramasser un mauvais champignon, toute votre récolte est foutue", vbExclamation, "Le voyageur"
                    Else
                        QtéChampignon = QtéChampignon + Qt    'Ajout des champignons dans le sac sinon
                    End If
                Else
                    QtéChampignon = QtéChampignon + Qt    'Ajout des champignons dans le sac
                End If
            ElseIf Picture1.Point(Image1.Left, Image1.Top) = 1310465 Then   'Si l'on est dans un bosquet à bon champignon
                Qt = Int((Rnd * 5) + 2) 'Nombre de Champignons ramasser
                MsgBox "Nombre de champignons ramasser : " & Qt, vbInformation, "Le voyageur"
                If SavoirChampi = False Then    'Test di l'on connais les champignons
                    If Int((Rnd * (Qt * 2)) + 1) = Qt Then 'Détermine si on tombe sur un mauvais champignon soit 1 chance sur 2 par champignons
                        MsgBox "Vous avez ramasser un mauvais champignon, toute votre récolte est foutue", vbExclamation, "Le voyageur"
                    Else
                        QtéChampignon = QtéChampignon + Qt    'Ajout des champignons dans le sac sinon
                    End If
                Else
                    QtéChampignon = QtéChampignon + Qt    'Ajout des champignons dans le sac
                End If
            ElseIf QtéFruit > 0 Or QtéChampignon > 0 Or QtéPoissons > 0 Then 'Si on est ailleur
                Manger.Show vbModal
            End If
    End Select
    X = Image1.Left + (Image1.Width / 2)
    Y = Image1.Top + (Image1.Height / 2)
    Select Case Picture1.Point(X, Y) 'Selon ou le smiley su trouve on définis le texte pour l'utilisateur
        Case 5056006
            Label1.Caption = "Vous êtes en pleine mer"
            Life = Life - 0.1   'Du à la fatigue
        Case 15956763   'Ok
            Label1.Caption = "Vous êtes au bord de la plage"
        Case 360821 'Ok
            Label1.Caption = "Vous êtes en raz campagne"
        Case 11285998   'Ok
            Label1.Caption = "Vous êtes au bord d'une falaise"
            Life = Life - 0.1 'Du à la fatigue
        Case 1310465    'Ok
            Label1.Caption = "Vous êtes dans un bosquet"
        Case 1080840    'Ok
            Label1.Caption = "Vous êtes en pleine forêt"
        Case 8846707    'Ok
            Label1.Caption = "Vous êtes sur un chemin forestier"
        Case 8126193    'Ok
            Label1.Caption = "Vous êtes à l'entrée d'une foret"
        Case 254
            Label1.Caption = "Vous êtes sur un chemin"
        Case 0  'Ok
            Label1.Caption = "Vous êtes sur un chemin de campagne"
        Case 16777215 'Si l'on est sur une zone blanche
            If PreviousColor = 16777215 Then Exit Sub
            If X >= 4980 And X <= 5250 And Y >= 855 And Y <= 1110 Then  'Détermine le marchand
                Vente.Show vbModal
            ElseIf (X >= 5390 And X <= 8265 And Y >= 6100 And Y <= 6630) Or (X >= 435 And Y >= 2520 And X <= 750 And Y <= 2820) Then    'Détermine le pêcheur
                If Nage = False Then    'Il faut d'abord savoir nager
                    Quest = MsgBox("Vous êtes chez un pêcheur, voulez-vous apprendre à nager ?" & Chr(13) & "Coût : 500", vbYesNo + vbQuestion, "Le voyageur")
                    If Quest = 6 Then
                        Nage = True
                        Somme = Somme - 500
                        faculte.AddItem "Nager comme un poisson"    'Ajout de la faculté
                        MsgBox "Vous pouvez maintenant apprendre à pêcher", vbInformation, "Le voyageur"
                    End If
                ElseIf Peche = False And Nage = True Then  'Il faut d'abord savoir nager
                    Quest = MsgBox("Vous êtes chez un pêcheur, voulez-vous apprendre à pêcher ?" & Chr(13) & "Coût : 500", vbYesNo + vbQuestion, "Le voyageur")
                    If Quest = 6 Then
                        Peche = True
                        Somme = Somme - 500
                        faculte.AddItem "Pêcher comme un requin"    'Ajout de la faculté
                    End If
                End If
            ElseIf X >= 4900 And X <= 6885 And Y >= 2600 And Y <= 4410 Then 'Détermine le druide
                If Vol = False Then 'Pour pouvoir savoir voler
                    Quest = MsgBox("Vous êtes chez un druide, voulez-vous apprendre à voler ?" & Chr(13) & "Couût : 700", vbYesNo + vbQuestion, "Le voyageur")
                    If Quest = 6 And Somme >= 700 Then
                        Vol = True
                        Somme = Somme - 700
                        faculte.AddItem "Voler comme un oiseau"
                    ElseIf Quest = 6 And Somme < 700 Then
                        MsgBox "Vente annulé, vous n'avez pas assez d'argent !!", vbExclamation, "Le voyageur"   'Pas de tunes, va bosser !!
                    End If
                End If
                If SavoirChampi = False Then 'Pour pouvoir savoir voler
                    Quest = MsgBox("Vous êtes chez un druide, voulez-vous apprendre à ceuillir les champignons ?" & Chr(13) & "Couût : 800", vbYesNo + vbQuestion, "Le voyageur")
                    If Quest = 6 And Somme >= 800 Then
                        SavoirChampi = True
                        Somme = Somme - 800
                        faculte.AddItem "Ceuillir les champignons"
                    ElseIf Quest = 6 And Somme < 800 Then
                        MsgBox "Vente annulé, vous n'avez pas assez d'argent !!", vbExclamation, "Le voyageur"   'Pas de tunes, va bosser !!
                    End If
                End If
            End If
    End Select
    PreviousColor = Picture1.Point(X, Y)
End Sub

 Private Sub Timer1_Timer()
    If Label1.Caption = "Vous êtes au bord d'une falaise" Then  'Si on est au bord d'une falaise
        If Vol = False Then
            Label2.Caption = "Vous ne savez pas voler, trouver un druide pour apprendre car vous tomber"
            Life = Life - 15    'Déduit de la vie au joueur, et oui il tombe !!
        End If
    ElseIf Label1.Caption = "Vous êtes au bord de la plage" Then    'Au bord de la plage
        If Nage = False Then Label2.Caption = "Vous devez trouver un pêcheur pour apprendre à nager"
    ElseIf Label1.Caption = "Vous êtes en pleine mer" Then  'En pleine mer
        If Nage = False Then
            Label2.Caption = "Vous êtes en train de vous noyer revenez au rivage"
            Life = Life - 10
        End If
    Else
        Label2.Caption = ""
    End If
    'Mise à jour des infos joueur
    Life = Life - 0.1
    Life = Format(Life, "### ##0.00")
    Label3.Caption = "Points de vie : " & Life
    Label4.Caption = "Nbr de fruits : " & QtéFruit
    Label6.Caption = "Nbr de champignons : " & QtéChampignon
    Label7.Caption = "Nbr de poissons : " & QtéPoissons
    Label5.Caption = "Somme : " & Somme
    If Life <= 0 Then
        MsgBox "Désoler, vous êtes décéder", vbCritical, "Le voyageur" 'Bhà t'es mort.
        newgame = MsgBox("Voulez-vous rejouer ?", vbYesNo + vbQuestion, "Le voyageur")
        If newgame = 6 Then
            Image1.Left = 7800
            Image1.Top = 5160
            Form_Load
        Else
            End
        End If
    End If
End Sub
