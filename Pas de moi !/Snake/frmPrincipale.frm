VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrincipale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   3450
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBete 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   2280
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog cmdCouleur 
      Left            =   1080
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl mmcMusique 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrCycle 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2520
      Top             =   240
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   -120
      X2              =   3960
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Image imgBete 
      Height          =   105
      Left            =   4320
      Picture         =   "frmPrincipale.frx":0000
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label lblTemps 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Image imgPomme 
      Height          =   105
      Left            =   4320
      Picture         =   "frmPrincipale.frx":005A
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label lblScoreAffichage 
      Caption         =   "Votre score est de:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu Joue 
         Caption         =   "Jouer"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Niv 
         Caption         =   "Niveau"
         Begin VB.Menu Fac 
            Caption         =   "Facile"
         End
         Begin VB.Menu Moy 
            Caption         =   "Moyen"
            Checked         =   -1  'True
         End
         Begin VB.Menu dif 
            Caption         =   "Difficile"
         End
      End
      Begin VB.Menu mnuCouleur 
         Caption         =   "Couleur du snake"
      End
      Begin VB.Menu mnuSupprimer 
         Caption         =   "Supprimer scores"
      End
      Begin VB.Menu aPropos 
         Caption         =   "A propos"
      End
      Begin VB.Menu ligne 
         Caption         =   "-"
      End
      Begin VB.Menu Quitte 
         Caption         =   "Quitter"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmPrincipale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************'
'Auteur: Sim�on Blanc                                                          '
'Lieu, Date: Saint-Croix, le 8 juin 2004                                       '
'Cr�er un jeux qui s'apparente au jeux du serpent sur t�l�phone portable       '
'******************************************************************************'

Option Explicit
'Le score est d�clar� en public, puisqu'il sera utilis� par deux feuilles
'(la feuille principale et la feuille des scores)
Public Score As Integer

'D�claration du tableau qui repr�sentera la feuille
Private tableau(0 To 39, 0 To 30) As Integer


'D�claration des coordonn�es de la fin et du d�but du serpent
Private premier_x, premier_y, dernier_x, dernier_y As Integer

'D�claration des coordonn�es de la pomme et de la b�te
Private pomme_x, pomme_y, direction, bete_x, bete_y As Integer

'D�claration de la variable qui indique si le serpent va �tre allong�
Private ajoutserpent As Boolean

'D�claration du nombre de point ajout� par pomme mang�e (d�pend de la difficult� choisie)
Private nbpointajoute As Integer

'D�claration de la couleur du serpent
Private couleur As String

'D�claration d'un compteur qui est incr�ment� de 1 � chaque fois que le joueur mange une pomme
'et quand le compteur arrive � 9 il affiche la b�te et remet le compteur � z�ro
Private compteurbete As Integer

'Constante pour les codes des touches qui sont mises dans la tableau d'entier
Const code_bas = 5
Const code_haut = 8
Const code_gauche = 4
Const code_droite = 6
Const code_pomme = 1
Const code_bete = 2

'Largeur de la case repr�sent�e sur la feuille
Const largcase = 100

'Largeur du tableau qui repr�sente la feuille
Const largeur = 39
'Hauteur du tableau qui repr�sente la feuille
Const hauteur = 30

'Vitesse pour les diff�rentes difficult�s
Const vitesse_facile = 40
Const vitesse_moyen = 25
Const vitesse_difficile = 15

'Nombre de points ajout�s par pomme suivant le niveau
Const point_facile = 6
Const point_moyen = 9
Const point_difficile = 12

Private Sub afficher()
    'Une "mise � jour" de la feuille est faite afin d'enlever les ancien rectangles
    'qui vont �tre remplac�s par les nouveaux
    frmPrincipale.Cls
    'Les deux boucles for avec un if, parcourt tout le tableau, pour afficher les celulles
    'marqu�es d'un autre chiffre que 1
    Dim d, i As Integer
    For i = 0 To largeur
        For d = 0 To hauteur
            If tableau(i, d) <> 0 Then
            
                'si la cellule = code de la pomme alors afficher un rectangle de couleur
                'diff�rente que le serpent
                If tableau(i, d) = code_pomme Then
                    'affichage
                    'Line (largcase * i, largcase * d)-(largcase * i + largcase, largcase * d + largcase), QBColor(2), BF
                    imgPomme.Move i * largcase, d * largcase
                Else
                    If tableau(i, d) = code_bete Then
                        imgBete.Move i * largcase, d * largcase
                    Else
                        'affichage du serpent
                        frmPrincipale.Line (largcase * i, largcase * d)-(largcase * i + largcase, largcase * d + largcase), couleur, BF
                    End If
                End If
            End If
            If i = premier_x And d = premier_y Then
                'Le FillStyle indique que le cercle qui va �tre dessin� va �tre plein
                FillStyle = 0
                'Dessin du cercle
                frmPrincipale.Circle ((i * largcase) + (largcase \ 2), (d * largcase) + (largcase \ 2)), largcase, QBColor(6)
            End If
        Next
    
    'Le programme g�re tous les �v�nements d'affichage avant de passer � la suite
    'DoEvents
    
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Ecris un code dans la variable direction si l'utilisateur tape une touche fl�ch�e
    'Pour que la touche soit enregistr�e, il ne faut pas que le serpent aille dans l'autre sens
    Select Case KeyCode
        Case vbKeyRight
            If direction <> code_gauche Then
                direction = code_droite
                'Debug.Print "DROITE"
            End If
    
        Case vbKeyLeft
            If direction <> code_droite Then
                direction = code_gauche
                'Debug.Print "GAUCHE"
            End If
    
        Case vbKeyDown
            If direction <> code_haut Then
                direction = code_bas
                'Debug.Print "BAS"
            End If
    
        Case vbKeyUp
            If direction <> code_bas Then
                direction = code_haut
                'Debug.Print "HAUT"
            End If
    End Select
    
    DoEvents
    
End Sub

Private Sub Joue_Click()
    Debug.Print "-- Initialisation des variables --"
    'Si la couleur = rien => la couleur n'a pas �t� choisie
    '=> couleur par d�faut
    If couleur = "" Then
        couleur = "1"
    End If
    
    'Remise � z�ro du score
    lblScore.Caption = "0"
    Score = 0

    Dim i, d As Integer
    'Mise a z�ro du tableau
    For i = 0 To largeur
        For d = 0 To hauteur
            tableau(i, d) = 0
        Next
    Next

    'Si le nombre de point � ajouter = 0 -> aucun changement
    '-> nb point par d�faut
    If nbpointajoute = 0 Then
        nbpointajoute = 9
    End If
    
    'Aucun ajout du serpent
    ajoutserpent = False
    
    'Ecriture du serpent de base dans le tableau
    For i = 2 To 7
        tableau(i, 12) = code_droite
    Next
    
    'Pomme visible
    imgPomme.Visible = True
    
    'Choix de la position de la pomme
    Debug.Print "-- G�n�ration d'un nouvel emplacement pour la pomme --"
    Do
    
        pomme_y = Int((hauteur - 1) * Rnd) + 1
        pomme_x = Int((largeur - 1) * Rnd) + 1
    
    Loop Until tableau(pomme_x, pomme_y) = 0
        
    tableau(pomme_x, pomme_y) = code_pomme
    'initialisation des variables de d�but et de fin du serpent
    premier_x = 7
    premier_y = 12
    dernier_x = 2
    dernier_y = 12
    'affiche d�s le d�but
    Call afficher
    'd�finit la direction de d�part
    direction = code_droite

    tmrCycle.Enabled = True

End Sub

Private Sub tmrCycle_Timer()

    On Error GoTo erreur

    'ce case augmente la t�te du serpent, en fonction de la direction choisie
    Select Case direction
        Case code_droite
            tableau(premier_x, premier_y) = code_droite
            premier_x = premier_x + 1
        Case code_gauche
            tableau(premier_x, premier_y) = code_gauche
            premier_x = premier_x - 1
        Case code_haut
            tableau(premier_x, premier_y) = code_haut
            premier_y = premier_y - 1
        Case code_bas
            tableau(premier_x, premier_y) = code_bas
            premier_y = premier_y + 1
    End Select

    If tableau(premier_x, premier_y) <> 0 And tableau(premier_x, premier_y) <> code_pomme And tableau(premier_x, premier_y) <> code_bete Then
        Debug.Print "-- Le serpent s'est touch� --"
        GoTo erreur
    End If
    
    If tableau(premier_x, premier_y) = code_bete Then
        lblScore = Val(lblScore.Caption) + Val(lblTemps.Caption)
        imgBete.Visible = False
        lblTemps.Caption = ""
        tmrBete.Enabled = False
    End If
 
    'Test afin de voir si la t�te du serpent a mang� une pomme
    If tableau(premier_x, premier_y) = code_pomme Then
    
        If compteurbete = 9 Then
            'Si 9 pommes ont �t� mang�es => faire appara�tre la b�te
            Debug.Print "-- Affichage de la b�te --"
            Call appBete
            compteurbete = 0
        End If

        compteurbete = compteurbete + 1

        ajoutserpent = True
        'ajout du score par le nb de point
        lblScore.Caption = Val(lblScore.Caption) + nbpointajoute
        
        'lancement de la musique de la pomme
        Debug.Print "-- Lancement du son de la pomme --"
        mmcMusique.Command = "Close"
        mmcMusique.DeviceType = "WaveAudio"
        mmcMusique.Notify = False
        mmcMusique.Wait = False
        mmcMusique.FileName = "pomme.wav"
        mmcMusique.Command = "Open"
        mmcMusique.Command = "Play"
        
        Do
            Randomize
            pomme_y = Int((hauteur - 1) * Rnd) + 1
            pomme_x = Int((largeur - 1) * Rnd) + 1
            Debug.Print "-- G�n�ration d'un nouvel emplacement pour la pomme --"
            'g�n�ration d'une nouvelle position de la pomme, jusqu'� ce que l'endroit de la pomme soit vide

        Loop Until tableau(pomme_x, pomme_y) = 0

        tableau(pomme_x, pomme_y) = code_pomme

    End If
    
    'ce case enl�ve des celulles � la queue du serpent, en fonction de la direction lorsque le cellule s'est inscrite
    If ajoutserpent = False Then
    
        Select Case tableau(dernier_x, dernier_y)
            Case code_droite
                tableau(dernier_x, dernier_y) = 0
                dernier_x = dernier_x + 1
            Case code_gauche
                tableau(dernier_x, dernier_y) = 0
                dernier_x = dernier_x - 1
            Case code_haut
                tableau(dernier_x, dernier_y) = 0
                dernier_y = dernier_y - 1
            Case code_bas
                tableau(dernier_x, dernier_y) = 0
                dernier_y = dernier_y + 1
        End Select
        
    Else
        Debug.Print "-- Ajout d'une cellule au serpent --"
        ajoutserpent = False
        
    End If
    
    Call afficher
'si une erreur est g�n�r�e, c'est le cas lorsque le serpent sort du tableau, le programme rentre
'dans cette proc�dure
Exit Sub

erreur:
    Debug.Print "-- Affichage du top10 --"
    'la feuille est cach�e et la feuille du top10 est montr�e
    tmrCycle.Enabled = False
    Score = Val(lblScore.Caption)
    frmPrincipale.Hide
    frmPerdu.Show
    lblTemps.Caption = ""
    tmrBete.Enabled = False
    imgBete.Visible = False
    
    imgPomme.Visible = False

End Sub

Sub appBete()

    Do
        Randomize
        bete_y = Int((hauteur - 1) * Rnd) + 1
        bete_x = Int((largeur - 1) * Rnd) + 1
        'g�n�ration d'une nouvelle position de la b�te, jusqu'� ce que l'endroit de la pomme soit vide
    Loop Until tableau(bete_y, bete_y) = 0

        tableau(bete_x, bete_y) = code_bete
        imgBete.Visible = True
    lblTemps.Caption = "10"
    tmrBete.Enabled = True
End Sub

Private Sub tmrBete_Timer()
    lblTemps.Caption = Val(lblTemps.Caption) - 1
    If Val(lblTemps.Caption) = 0 Then
        lblTemps.Caption = ""
        tmrBete.Enabled = False
        imgBete.Visible = False
        Debug.Print "-- Effacage de la b�te --"
    End If
End Sub

Private Sub dif_Click()
    
    Debug.Print "-- Mise du niveau � Difficile --"
    
    'Rapidit� de la b�te � dispara�tre
    tmrBete.Interval = 200
    
    'Le niveau du jeu est mis a "difficile", le nombre de points par pomme est � 12
    nbpointajoute = point_difficile
    
    'Les flag (ou coches) dans les menu permetent de selectionner le niveau choisi
    dif.Checked = True
    Moy.Checked = False
    Fac.Checked = False
    
    'La difficult� du serpent est initialis�e � difficile
    tmrCycle.Interval = vitesse_difficile
End Sub

Private Sub Moy_Click()

    Debug.Print "-- Mise du niveau � Moyen --"
    
   'Rapidit� de la b�te � dispara�tre
    tmrBete.Interval = 300
    
    'Le niveau du jeu est mis a "moyen", le nombre de points par pomme est � 9
    nbpointajoute = point_moyen
    
    'Les flag (ou coches) dans les menu permetent de selectionner le niveau choisi
    dif.Checked = False
    Moy.Checked = True
    Fac.Checked = False
    
    'La difficult� du serpent est initialis�e � moyen
    tmrCycle.Interval = vitesse_moyen
    
End Sub

Private Sub Fac_Click()

    Debug.Print "-- Mise du niveau � Facile --"

   'Rapidit� de la b�te � dispara�tre
    tmrBete.Interval = 500
    
    'Le niveau du jeu est mis a "facile", le nombre de points par pomme est � 6
    nbpointajoute = point_facile
    
    'Les flag (ou coches) dans les menu permetent de selectionner le niveau choisi
    dif.Checked = False
    Moy.Checked = False
    Fac.Checked = True
    
    'La difficult� du serpent est initialis�e � facile
    tmrCycle.Interval = vitesse_facile
    
End Sub

Private Sub aPropos_Click()
    'Montre la feuille d'A propos
    frmAPropos.Show
End Sub


Private Sub mnuCouleur_Click()

    Debug.Print "-- Changement de couleur du serpent --"
    'Afficher la bo�te de dialogue pour le choix d'une couleur
    cmdCouleur.ShowColor
    
    'Mise du code hexad�cimal dans la variable couleur
    couleur = cmdCouleur.Color

End Sub

Private Sub mnuSupprimer_Click()
    Debug.Print "-- Supression du fichier des r�sultats --"
    'Donne le chmin d'un fichier qui se situe "a c�t�" de l'executable
    Dim NomFichier, chemin As String
    chemin = CurDir & "\" & "resultat.snk"
    'Si le fichier existe
    If Dir(chemin) = "resultat.snk" Then
        Debug.Print "=> Le fichier existait => il a �t� effac�"
        'l'�liminer
        Kill (chemin)
    Else
        'sinon afficher qu'il a d�j� �t� effac�
        Debug.Print "=> Le fichier n'existe pas"
        MsgBox "Le fichier a d�j� �t� effac�", vbCritical, "Message d'erreur"
    End If
        
End Sub

Private Sub Quitte_Click()
    End
End Sub

Public Sub retourtop10()
    Debug.Print "-- Initialisation des variables pour retourner sur la feuille --"
    'Mise a z�ro du tableau
    Dim i, d As Integer
    For i = 0 To largeur
        For d = 0 To hauteur
            tableau(i, d) = 0
        Next
    Next
    
    'Mise du score � z�ro
    lblScore.Caption = 0
    
    compteurbete = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

