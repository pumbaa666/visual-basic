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

'D�claration du type qui est stock�e dans un fichier texte
Private Type classement

    'Pseudonyme de la personne qui a jou�
    Pseudo As String * 50
    
    'Score de la personne qui a jou�
    Score As String * 50
 
End Type

'D�claration du top10 comme classement
Private top10 As classement


Private Sub cmdRetoure_Click()

    Debug.Print "-- Retour sur la feuille --"
    
    'La forme principale (de jeu) est affich�e
    frmPrincipale.Show
    
    'La forme du top10 est cach�e
    frmPerdu.Hide
    
    'La proc�dure d'initialisation de la feuille principale (de jeu) est lanc�e afin que l'utilisateur
    'puisse jouer directement
    Call frmPrincipale.retourtop10

End Sub

Public Sub Form_Activate()

    Debug.Print "-- Ecriture du score dans le caption de la feuille --"
    
    'Le score de l'utilisateur est affich� dans l'ent�te de la feuille du top10
    frmPerdu.Caption = "Votre score est de: " & frmPrincipale.Score

    'le chemin est cr�� pour que le fichier texte annexe soit au meme endroit que le programme
    chemin = CurDir & "\resultat.snk"
    
    Debug.Print "-- Ouverture du fichier --"
    
    'Ouverture du fichier en acc�s s�quentiel
    Open chemin For Random As #1 Len = Len(top10)

    'compare le r�sultat obtenu avec tout les r�sultat du top10
    'si le compteur <> 0, alors c'est que le joueur n'est pas dans le top10
    'sinon, la variable compteur indique le r�sultat
    Dim i As Integer
    
    'D�claration et initialisation de la variable qui contient
    'rang de la personne qui a jou�
    Dim resultat As Integer
    
    Debug.Print "-- Recherche du r�sultat du joueur --"
    
    'La boucle for parcour tout le top10 pour sauvegarder la place du joueur
    For i = 1 To 9
    
        'Prise des informations dans le fichier � l'enregistrement i
        Get #1, i, top10
        
        'Compare si l'utilisateur a fait un meilleur score que la personne qui a l'enregistrement i
        If Val(top10.Score) < frmPrincipale.Score Or top10.Score = "" Then

            'Si la personne avait d�j� un rang qui �tait calcul� (resultat <> 0), il ne faut pas
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
        
        Debug.Print "-- D�calage des r�sultats --"
        
        'D�caler les r�sultat deri�re lui afin que personne ne soit effac�
        
        'D�claration des variables pour faire un �change de place afin que tous le monde soit d�cal�
        Dim trans1, trans2, trans3, trans4 As String
        
        'Variable qui permet de faire une alternance entre deux possibilit�s
        Dim code As Integer
    
        'La variable de d�but d'enregistrement est le rang du joueur
        i = resultat

        'Sauvgarde du joueur qui avait le rang deri�re la personnes
        Get #1, i + 1, top10
        trans1 = top10.Score
        trans3 = top10.Pseudo
        
        'D�calage de la personne � laquelle on a pris son rang
        Get #1, i, top10
        'D�calage de (i + 1)
        Put #1, i + 1, top10

        'La boucle for parcourt tout le classement pour d�caler tous le monde afin que personnes ne soit l�s�
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
        top10.Pseudo = InputBox("Donnez votre pseudo", "Mise en m�moire du pseudo")
        
        'Il �crit les donn�es dans le fichier texte
        Put #1, resultat, top10
    
    End If
    
    'Le case lance certains sons qui sont choisis suivant la place du joueur
    Select Case resultat
    
    Case 1:
    
        'S'il est premier un son sp�cial est jou�
        Debug.Print "-- Son pour le premier r�sultat --"
        frmPrincipale.mmcMusique.Command = "Close"
        frmPrincipale.mmcMusique.DeviceType = "WaveAudio"
        frmPrincipale.mmcMusique.Notify = False
        frmPrincipale.mmcMusique.Wait = False
        frmPrincipale.mmcMusique.FileName = "premier.wav"
        frmPrincipale.mmcMusique.Command = "Open"
        frmPrincipale.mmcMusique.Command = "Play"
        
    Case 0:

        'S'il n'est pas dans le top10 un son sp�cial est jou�
        Debug.Print "-- Son pour le r�sultat qui ne rentre pas dans le top 10 --"
        frmPrincipale.mmcMusique.Command = "Close"
        frmPrincipale.mmcMusique.DeviceType = "WaveAudio"
        frmPrincipale.mmcMusique.Notify = False
        frmPrincipale.mmcMusique.Wait = False
        frmPrincipale.mmcMusique.FileName = "perdu.wav"
        frmPrincipale.mmcMusique.Command = "Open"
        frmPrincipale.mmcMusique.Command = "Play"
        
    Case Else:
        
        'S'il n'est pas premier et dans le top10 un son est jou�
        Debug.Print "-- Son pour le r�sultat ins�r� dans le top 10 --"
        frmPrincipale.mmcMusique.Command = "Close"
        frmPrincipale.mmcMusique.DeviceType = "WaveAudio"
        frmPrincipale.mmcMusique.Notify = False
        frmPrincipale.mmcMusique.Wait = False
        frmPrincipale.mmcMusique.FileName = "top10.wav"
        frmPrincipale.mmcMusique.Command = "Open"
        frmPrincipale.mmcMusique.Command = "Play"

    End Select

    Debug.Print "-- Affichage des r�sultat sur la feuille --"
    
    'La boucle for parcour tout le fichier des scores et affiche les score ainsi que les pseudo dans les labels pr�vus
    'a cet effet sur la feuille (les labels ont un index pour ne pas faire un nom sp�cial pour chacun d'eux)
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

