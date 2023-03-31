Attribute VB_Name = "Module1"
'formatage de l'ent�te d'un fichier Bitmap
Type EnteteFichierBmp
    EFBFileType As String * 2   '� BM � : signature du fichier
    EFBFileSize As Long         'Taille du fichier en octets
    EFBReserved As Long         'Toujours � 0 : r�serv�
    EFBBitMapOffset As Long     'Offset de l�image � partir du d�but du fichier
    EFBHeaderSize As Long       'Taille de l�en-t�te en octets
    EFBWidth As Long            'Largeur en pixel de l�image
    EFBHeight As Long           'Hauteur en pixel de l�image
                        'Si Height prend une valeur n�gative
                        '(ce qui est tr�s rare),
                        'c�est que l�image n�est pas retourn�e
    EFBPlanes As Integer        'Nombre de plans utilis�s (normalement � 1)
    EFBBitsPerPixel As Integer  'Nombre de bits par pixel
    EFBCompression As Long      'M�thode de compression
    EFBSizeOfBitMap As Long     'Taille de l�image en octets
    EFBHorzResolution As Long   'R�solution horizontale en pixels
    EFBVertResolution As Long   'R�solution verticale en pixel
    EFBColorsUsed As Long       'Nombre de couleur dans la palette
                                'Si  0 : palette enti�re utilis�e
    EFBColorsImportant As Long  'Nombre de couleurs importantes
    
End Type

'sert � lire et � remplir le corps d'un fichier bitmap en niveaux de gris
'ainsi qu'� remplir la palette du st�r�ogramme
Type CorpsFichierBmp
    CouleurPixel As Byte
End Type

'pour la lecture d'un fichier palette
Type FichierPal
    PalR As Byte
    PalG As Byte
    PalB As Byte
End Type

'variables de sauvegarde des caract�ristiques utiles de la texture en cours
'(utilis� pour la cr�ation du st�r�ogramme)
Public TexHResol As Long
Public TexVResol As Long
Public PaletteTexD3D(256, 4) As Byte
Public TexCorps() As Byte

'variables utilis�es lors de la cr�ation du st�r�ogramme
Public CX As Long
Public CY As Long
Public HResol As Long
Public VResol As Long
Public NbBnd As Long

'tableau utilis� pour les calculs utiles � la cr�ation du st�r�ogramme
Public Couleur() As Long
'tableaux utilis� pour la sauvegarde de la palette courante
Public PaletteD3D(256, 4) As Byte
Public SavPaletteD3D(256, 4) As Byte


'r�glages sauvegard�s dans reglages.ini
'(langue utilis�e, largeur d'une bande, nombre de profondeurs maximum, flag,
'nombre de couleurs utilis�es dans la palette, utilisation d'une palette de
'couleurs ou d'une image pour cr�er le st�r�ogramme)
Public Lng As String
Public LrgBnd As Long
Public NbProf As Long
Public FctNbProf As Long
Public NbCoulPal As Long
Public Resultat3D As String

Public FormActive As String

'tableaux utilis�s pour sauvegarder les r�glages et les "captions"
Public TabLng(50) As String
Public TabRg(50) As String

'quelques fichiers
Public FichierAide As String
Public FichierNomLongAide As String
Public FichierNomLongTexture As String
Public FichierNomLongPalette As String

