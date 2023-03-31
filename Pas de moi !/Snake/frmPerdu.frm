VERSION 5.00
Begin VB.Form frmPerdu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vous avez perdu :-)"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetoure 
      Caption         =   "Retour"
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblTop10 
      Alignment       =   2  'Center
      Caption         =   "TOP 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   19
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblAfficheScore 
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblAfficheNom 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmPerdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Le chemin du fichier qui stocke les scores
Dim chemin As String

'Déclaration du type qui est stockée dans un fichier texte
Private Type classement

    'Pseudonyme de la personne qui a joué
    Pseudo As String * 50
    
    'Score de la personne qui a joué
    Score As String * 50
 
End Type

'Déclaration du top10 comme classement
Private top10 As classement


Private Sub cmdRetoure_Click()

    Debug.Print "-- Retour sur la feuille --"
    
    'La forme principale (de jeu) est affichée
    frmPrincipale.Show
    
    'La forme du top10 est cachée
    frmPerdu.Hide
    
    'La procédure d'initialisation de la feuille principale (de jeu) est lancée afin que l'utilisateur
    'puisse jouer directement
    Call frmPrincipale.retourtop10

End Sub

Public Sub Form_Activate()

    Debug.Print "-- Ecriture du score dans le caption de la feuille --"
    
    'Le score de l'utilisateur est affiché dans l'entête de la feuille du top10
    frmPerdu.Caption = "Votre score est de: " & frmPrincipale.Score

    'le chemin est créé pour que le fichier texte annexe soit au meme endroit que le programme
    chemin = CurDir & "\resultat.snk"
    
    Debug.Print "-- Ouverture du fichier --"
    
    'Ouverture du fichier en accès séquentiel
    Open chemin For Random As #1 Len = Len(top10)

    'compare le résultat obtenu avec tout les résultat du top10
    'si le compteur <> 0, alors c'est que le joueur n'est pas dans le top10
    'sinon, la variable compteur indique le résultat
    Dim i As Integer
    
    'Déclaration et initialisation de la variable qui contient
    'rang de la personne qui a joué
    Dim resultat As Integer
    
    Debug.Print "-- Recherche du résultat du joueur --"
    
    'La boucle for parcour tout le top10 pour sauvegarder la place du joueur
    For i = 1 To 9
    
        'Prise des informations dans le fichier à l'enregistrement i
        Get #1, i, top10
        
        'Compare si l'utilisateur a fait un meilleur score que la personne qui a l'enregistrement i
        If Val(top10.Score) < frmPrincipale.Score Or top10.Score = "" Then

            'Si la personne avait déjà un rang qui était calculé (resultat <> 0), il ne faut pas
            'lui donner un moins bon resultat
            If resultat = 0 Then
                resultat = i
            End If

        End If
        
    Next
    
    
    'Si l'utilisateur n'est pas dans le top 10
    If resultat = 0 Then
    
        'Afficher qu'il n'y est pas
        MsgBox "Vous ne rentrez pas dans le TOP10!", vbInformation
        
    Else
        
        Debug.Print "-- Décalage des résultats --"
        
        'Décaler les résultat derière lui afin que personne ne soit effacé
        
        'Déclaration des variables pour faire un échange de place afin que tous le monde soit décalé
        Dim trans1, trans2, trans3, trans4 As String
        
        'Variable qui permet de faire une alternance entre deux possibilités
        Dim code As Integer
    
        'La variable de début d'enregistrement est le rang du joueur
        i = resultat

        'Sauvgarde du joueur qui avait le rang derière la personnes
        Get #1, i + 1, top10
        trans1 = top10.Score
        trans3 = top10.Pseudo
        
        'Décalage de la personne à laquelle on a pris son rang
        Get #1, i, top10
        'Décalage de (i + 1)
        Put #1, i + 1, top10

        'La boucle for parcourt tout le classement pour décaler tous le monde afin que personnes ne soit lésé
        For i = resultat + 1 To 9

            If code = 0 Then
            
                'Sauvgarde de l'enregistrement (i+1)
                Get #1, i + 1, top10
                trans2 = top10.Score
                trans4 = top10.Pseudo
                
                'Ecriture du joueur devant dans la place actuelle
                top10.Score = trans1
                top10.Pseudo = trans3
                Put #1, i + 1, top10
                
                'Prochain passage dans l'autre partie du if
                code = 1
        
            Else
            
                'Sauvegarde de l'enregistrement (i+1)
                Get #1, i + 1, top10
                trans1 = top10.Score
                trans3 = top10.Pseudo
                
                'Ecriture du joueur devant dans la place actuelle
                top10.Pseudo = trans4
                top10.Score = trans2
                Put #1, i + 1, top10

                'Prochain passage dans l'autre partie du if
                code = 0
                
            End If

        Next

    End If
    
    'Si le joueur fait partie du top10
    If resultat <> 0 Then
    
        'Enregistrement du score sur le rang du joueur
        top10.Score = frmPrincipale.Score
        
        Debug.Print "-- Enregistrement du serpent --"
        
        'Si le compteur est dans le top10, le programme demande le pseudo du joueur
        top10.Pseudo = InputBox("Donnez votre pseudo", "Mise en mémoire du pseudo")
        
        'Il écrit les données dans le fichier texte
        Put #1, resultat, top10
    
    End If
    
    'Le case lance certains sons qui sont choisis suivant la place du joueur
    Select Case resultat
    
    Case 1:
    
        'S'il est premier un son spécial est joué
        Debug.Print "-- Son pour le premier résultat --"
        frmPrincipale.mmcMusique.Command = "Close"
        frmPrincipale.mmcMusique.DeviceType = "WaveAudio"
        frmPrincipale.mmcMusique.Notify = False
        frmPrincipale.mmcMusique.Wait = False
        frmPrincipale.mmcMusique.FileName = "premier.wav"
        frmPrincipale.mmcMusique.Command = "Open"
        frmPrincipale.mmcMusique.Command = "Play"
        
    Case 0:

        'S'il n'est pas dans le top10 un son spécial est joué
        Debug.Print "-- Son pour le résultat qui ne rentre pas dans le top 10 --"
        frmPrincipale.mmcMusique.Command = "Close"
        frmPrincipale.mmcMusique.DeviceType = "WaveAudio"
        frmPrincipale.mmcMusique.Notify = False
        frmPrincipale.mmcMusique.Wait = False
        frmPrincipale.mmcMusique.FileName = "perdu.wav"
        frmPrincipale.mmcMusique.Command = "Open"
        frmPrincipale.mmcMusique.Command = "Play"
        
    Case Else:
        
        'S'il n'est pas premier et dans le top10 un son est joué
        Debug.Print "-- Son pour le résultat inséré dans le top 10 --"
        frmPrincipale.mmcMusique.Command = "Close"
        frmPrincipale.mmcMusique.DeviceType = "WaveAudio"
        frmPrincipale.mmcMusique.Notify = False
        frmPrincipale.mmcMusique.Wait = False
        frmPrincipale.mmcMusique.FileName = "top10.wav"
        frmPrincipale.mmcMusique.Command = "Open"
        frmPrincipale.mmcMusique.Command = "Play"

    End Select

    Debug.Print "-- Affichage des résultat sur la feuille --"
    
    'La boucle for parcour tout le fichier des scores et affiche les score ainsi que les pseudo dans les labels prévus
    'a cet effet sur la feuille (les labels ont un index pour ne pas faire un nom spécial pour chacun d'eux)
    For i = 1 To 9
    
        'Va chercher l'enregistrement parcouru par la boucle for avec comme No d'enregistrement i
        Get #1, i, top10
        
        'Et affichage de l'enregistrement sur la feuille
        lblAfficheNom(i - 1).Caption = top10.Pseudo
        lblAfficheScore(i - 1).Caption = top10.Score

    Next
    
    'fermeture du fichier
    Close #1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

