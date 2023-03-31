Attribute VB_Name = "Module1"
'formatage de l'entête d'un fichier Bitmap
Type EnteteFichierBmp
    EFBFileType As String * 2   '« BM » : signature du fichier
    EFBFileSize As Long         'Taille du fichier en octets
    EFBReserved As Long         'Toujours à 0 : réservé
    EFBBitMapOffset As Long     'Offset de l’image à partir du début du fichier
    EFBHeaderSize As Long       'Taille de l’en-tête en octets
    EFBWidth As Long            'Largeur en pixel de l’image
    EFBHeight As Long           'Hauteur en pixel de l’image
                        'Si Height prend une valeur négative
                        '(ce qui est très rare),
                        'c’est que l’image n’est pas retournée
    EFBPlanes As Integer        'Nombre de plans utilisés (normalement à 1)
    EFBBitsPerPixel As Integer  'Nombre de bits par pixel
    EFBCompression As Long      'Méthode de compression
    EFBSizeOfBitMap As Long     'Taille de l’image en octets
    EFBHorzResolution As Long   'Résolution horizontale en pixels
    EFBVertResolution As Long   'Résolution verticale en pixel
    EFBColorsUsed As Long       'Nombre de couleur dans la palette
                                'Si  0 : palette entière utilisée
    EFBColorsImportant As Long  'Nombre de couleurs importantes
    
End Type

'sert à lire et à remplir le corps d'un fichier bitmap en niveaux de gris
'ainsi qu'à remplir la palette du stéréogramme
Type CorpsFichierBmp
    CouleurPixel As Byte
End Type

'pour la lecture d'un fichier palette
Type FichierPal
    PalR As Byte
    PalG As Byte
    PalB As Byte
End Type

'variables de sauvegarde des caractéristiques utiles de la texture en cours
'(utilisé pour la création du stéréogramme)
Public TexHResol As Long
Public TexVResol As Long
Public PaletteTexD3D(256, 4) As Byte
Public TexCorps() As Byte

'variables utilisées lors de la création du stéréogramme
Public CX As Long
Public CY As Long
Public HResol As Long
Public VResol As Long
Public NbBnd As Long

'tableau utilisé pour les calculs utiles à la création du stéréogramme
Public Couleur() As Long
'tableaux utilisé pour la sauvegarde de la palette courante
Public PaletteD3D(256, 4) As Byte
Public SavPaletteD3D(256, 4) As Byte


'réglages sauvegardés dans reglages.ini
'(langue utilisée, largeur d'une bande, nombre de profondeurs maximum, flag,
'nombre de couleurs utilisées dans la palette, utilisation d'une palette de
'couleurs ou d'une image pour créer le stéréogramme)
Public Lng As String
Public LrgBnd As Long
Public NbProf As Long
Public FctNbProf As Long
Public NbCoulPal As Long
Public Resultat3D As String

Public FormActive As String

'tableaux utilisés pour sauvegarder les réglages et les "captions"
Public TabLng(50) As String
Public TabRg(50) As String

'quelques fichiers
Public FichierAide As String
Public FichierNomLongAide As String
Public FichierNomLongTexture As String
Public FichierNomLongPalette As String

