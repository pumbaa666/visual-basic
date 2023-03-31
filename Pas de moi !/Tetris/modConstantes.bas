Attribute VB_Name = "modConstantes"
Option Explicit

' audio
Public Enum EWave
    WAV_TOC = 1
    WAV_NIVEAU = 2
    WAV_LIGNE = 3
    WAV_MOUVEMENT = 4
End Enum
Public Enum EMidi
    MID_TETRIS = 1
    MID_FIN = 2
End Enum

' graphique
Public Enum ECouleur
    VIDE = 0
    ROUGE = 1
    BLEU = 2
    JAUNE = 3
    GRIS = 4
    VERT = 5
    NOIR = 6                            ' seulement pour les lignes complètes
End Enum

' jeu
Public Enum EStyleJeu
    JEU_LIGNE = 1
    JEU_COULEUR = 2
End Enum
Public Enum EMouvement
    GAUCHE = 1
    DROITE = 2
    ROT_POS = 3
    BAS = 4
    ROT_NEG = 5
    HAUT = 6
End Enum

Public Const TAG_JEU = "Jeu"            ' tag des timers
Public Const TAG_ANIM = "Animation"

Public Const MAX_X = 10, MAX_Y = 20     ' dimensions du jeu

Public Const MAX_COULEUR = 5            ' nombre de couleurs des pièces

Public Const TEMPO_INI = 15             ' tempo de la musique (lent)
Public Const TEMPO_FIN = 5              ' tempo rapide
Public Const TEMPO_NORMAL = 6           ' tempo par défaut

Public Const DUREE_CHUTE_LIBRE = 1      ' vitesse de chute de la pièce
Public Const DUREE_CHUTE_INI = 150      ' descente automatique (1er niveau)
Public Const DUREE_CHUTE_FIN = 5        ' descente automatique (dernier niveau)
Public Const INTERVALLE_JEU = 10        ' timer joueur
Public Const INTERVALLE_ANIM = 180      ' timer animation

Public Const MAX_ANIMATION = 7          ' nombre de clignotements

Public Const MAX_NIVEAU = 9
Public Const LIGNE_NIVEAU = 10          ' nbr de lignes pour changer de niveau
Public Const BLOC_NIVEAU = 40           ' nbr de blocs pour changer de niveau

' nbr minimal de blocs de la même couleur pour qu'ils soient supprimés
Public Const MIN_VOISINS = 4
