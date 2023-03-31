Attribute VB_Name = "Moteur"
'  ******************************************************************
'  *                                                                *
'  *       Création du stéréogramme en Bitmap à partir              *
'  *       d'un Bitmap en niveaux de gris                           *
'  *                                                                *
'  ******************************************************************

Sub CreationBmpNivGris()

Dim UdtEnrFichier As EnteteFichierBmp
Dim UdtEnrCorpsFichier As CorpsFichierBmp
Dim filecal
Dim Fichier As String
Dim FichierD3D As String
Dim Indice As String
Dim Avanc As Long
Dim decalageX As Long
Dim CX2 As Long
Dim CX3 As Long
Dim TestTaille As Integer
Dim TestTailleF As Integer
Dim Position As Long

'----------------------------------------------------------------'
'    ***  chargement du bitmap dans tableau de traitement  ***   '
'----------------------------------------------------------------'
    'si non de fichier correspondant à la fenêtre active non renseigné
    'alors inutile de continuer
    If FormActive = "" Then
        Exit Sub
    End If
    Open FormActive For Random As #1 Len = Len(UdtEnrFichier)
        'si fichier vide...
        If Len(FormActive) = 0 Then
            Exit Sub
        End If
    FichierD3D = FormActive + ".D3D.bmp"
    Get #1, 1, UdtEnrFichier

        NbBnd = Int(CX / LrgBnd)
        HResol = UdtEnrFichier.EFBWidth
        TestTailleF = HResol
        VResol = UdtEnrFichier.EFBHeight
        CX = UdtEnrFichier.EFBWidth + NbProf
        CY = UdtEnrFichier.EFBHeight
    Close #1
'calcul de la taille réelle du fichier (hresol)
If TestTailleF Mod 2 = 0 Then
    If TestTailleF Mod 4 = 0 Then
        TestTailleF = 0
    Else
        TestTailleF = 2
        HResol = HResol + TestTailleF
    End If
Else
    TestTailleF = TestTailleF + 1
    If TestTailleF Mod 4 = 0 Then
        TestTailleF = 1
        HResol = HResol + TestTailleF
    Else
        TestTailleF = 3
        HResol = HResol + TestTailleF
    End If
End If

'affiche boite dialogue attente
Avancement.Pgbar.Value = 0
Avancement.Show
Avancement.Pgbar.Max = VResol
Avancement.phase1(0).Caption = "Chargement des données"
Avancement.phase1(1).Caption = ""
Avancement.phase1(2).Caption = ""

'chargement du tableau "couleur(1=profondeur d'origine
'et 2=profondeur finale et 3=origine du point,Cy,Cx)" avec le fichier bmp
ReDim Couleur(4, VResol, HResol + NbProf + LrgBnd + 2)
Open FormActive For Random As #1 Len = Len(UdtEnrCorpsFichier)
Position = UdtEnrFichier.EFBBitMapOffset
For CY = VResol To 1 Step -1
    For CX = 1 To HResol
        Position = Position + 1
        Get #1, Position, UdtEnrCorpsFichier
        Couleur(1, CY, CX) = UdtEnrCorpsFichier.CouleurPixel
    Next
    Avancement.Pgbar.Value = Avancement.Pgbar.Value + 1
    Avanc = Int(Avancement.Pgbar.Value / VResol * 100)
    Avancement.lbl.Caption = Avanc
    DoEvents
Next
Close #1


'boite dialogue attente
Avancement.Pgbar.Value = 0
Avancement.Pgbar.Max = VResol
Avancement.phase1(0).Caption = "Données chargées"
Avancement.phase1(1).Caption = "Calculs en cours"

'----------------------------------------------------------------'
'    ***  phase de calcul des décalages à réaliser  ***          '
'----------------------------------------------------------------'
'couleur(1,y,x) = couleur d'origine
'couleur(2,y,x) = profondeur finale calculé en fonction de la couleur d'origine
'couleur(3,y,x) = origine du point
'couleur(4,y,x) = couleur du dessin final

CX3 = HResol + NbProf + LrgBnd

'calcul des décalages à faire dans 2 et 3
For CY = 1 To VResol
    For CX = 1 To HResol - TestTailleF
        profondeur = NbProf + 1 - Int(Couleur(1, CY, CX) / FctNbProf)
        'profondeur = NbProf - Int(Couleur(1, CY, CX) / FctNbProf) - 1
        CX2 = CX + profondeur + LrgBnd
        If Couleur(2, CY, CX2) < profondeur Then
            Couleur(2, CY, CX2) = profondeur
            Couleur(3, CY, CX2) = CX
        End If
    Next

    'si on met un fond random à partir d'une palette donnée
    If Resultat3D = "Random" Then
        'initialisation du dessin de la première bande dans 4
        For CX = 1 To LrgBnd
            c = CInt(Int(Rnd * NbCoulPal))
            Couleur(4, CY, CX) = c
        Next
        'calcul du dessin des autres bandes
        For CX = LrgBnd + 1 To CX3
            If Couleur(2, CY, CX) = 0 Then
                c = CInt(Int(Rnd * NbCoulPal))
                Couleur(4, CY, CX) = c
            Else
                Couleur(4, CY, CX) = Couleur(4, CY, Couleur(3, CY, CX))
            End If
        Next
    End If
    'si on met un fond Texture à partir d'une palette donnée
    If Resultat3D = "Texture" Then
        'initialisation du dessin de la première bande dans 4
        TY = (CY Mod TexVResol) + 1
        For CX = 1 To LrgBnd
            TX = CX Mod TexHResol
            c = TexCorps(TY, TX)
            Couleur(4, CY, CX) = c
        Next
        'calcul du dessin des autres bandes
        For CX = LrgBnd + 1 To CX3
            If Couleur(2, CY, CX) = 0 Then
                'TX = CX Mod TexHResol
                'c = TexCorps(TY, TX)
                'Couleur(4, CY, CX) = c
                Couleur(4, CY, CX) = Couleur(4, CY, CX - LrgBnd + 1)
            Else
                Couleur(4, CY, CX) = Couleur(4, CY, Couleur(3, CY, CX))
            End If
        Next
    End If
    
    Avancement.Pgbar.Value = Avancement.Pgbar.Value + 1
    Avanc = Int(Avancement.Pgbar.Value / VResol * 100)
    Avancement.lbl.Caption = Avanc
    DoEvents
Next

'----------------------------------------------------------------'
'    ***  création du fichier bitmap temporaire  ***             '
'----------------------------------------------------------------'
Avancement.Pgbar.Value = 0
Avancement.Pgbar.Max = VResol
Avancement.phase1(1).Caption = "Calculs terminés"
Avancement.phase1(2).Caption = "Création du fichier Bitmap Stéréogramme"
'creation fichiers bmp résultant
'création entête
filecal = 77
Open FichierD3D _
For Random As filecal Len = Len(UdtEnrFichier)
    UdtEnrFichier.EFBFileType = "BM"
    UdtEnrFichier.EFBFileSize = CX3 * VResol + 1078
    UdtEnrFichier.EFBReserved = 0
    UdtEnrFichier.EFBBitMapOffset = 54 + 256 * 4    '256
    UdtEnrFichier.EFBHeaderSize = 40
    UdtEnrFichier.EFBWidth = CX3
    UdtEnrFichier.EFBHeight = VResol
    UdtEnrFichier.EFBPlanes = 1
    UdtEnrFichier.EFBBitsPerPixel = 8
    UdtEnrFichier.EFBCompression = 0
    UdtEnrFichier.EFBSizeOfBitMap = CX3 * VResol
    UdtEnrFichier.EFBHorzResolution = 5830
    UdtEnrFichier.EFBVertResolution = 5830
    UdtEnrFichier.EFBColorsUsed = 0
    UdtEnrFichier.EFBColorsImportant = 0
Put filecal, 1, UdtEnrFichier
Close filecal

'création de la palette de couleur
Open FichierD3D _
For Random As filecal Len = Len(UdtEnrCorpsFichier)
Indice = 54
If Resultat3D = "Random" Then
    For I = 0 To 255    '255
        For J = 1 To 4
            Indice = Indice + 1
            UdtEnrCorpsFichier.CouleurPixel = PaletteD3D(I, J)
            Put filecal, Indice, UdtEnrCorpsFichier
        Next
    Next
End If
If Resultat3D = "Texture" Then
    For I = 0 To 255    '255
        For J = 1 To 4
            Indice = Indice + 1
            UdtEnrCorpsFichier.CouleurPixel = PaletteTexD3D(I, J)
            Put filecal, Indice, UdtEnrCorpsFichier
        Next
    Next
End If

'création des pixels
'Indice = 1078

'si 1,2... alors +2

'si 3,4... alors +0
TestTaille = LrgBnd
If TestTaille Mod 2 <> 0 Then
    TestTaille = TestTaille + 1
End If
TestTaille = TestTaille / 2
If TestTaille Mod 2 <> 0 Then
    TestTaille = 2
Else
    TestTaille = 0
End If
For CY = VResol To 1 Step -1
'TestTaille sert à mettre un nombre d'octets pair dans une ligne de l'image
    For CX = 1 To CX3 + TestTaille
        Indice = Indice + 1
'        UdtEnrCorpsFichier.CouleurPixel = Couleur(1, CY, CX)
        UdtEnrCorpsFichier.CouleurPixel = Couleur(4, CY, CX)
        Put filecal, Indice, UdtEnrCorpsFichier
    Next
    If CX3 / 2 - Int(CX3 / 2) <> 0 Then
        Indice = Indice + 1
        UdtEnrCorpsFichier.CouleurPixel = 0
        Put filecal, Indice, UdtEnrCorpsFichier
    End If
    Avancement.Pgbar.Value = Avancement.Pgbar.Value + 1
    Avanc = Int(Avancement.Pgbar.Value / VResol * 100)
    Avancement.lbl.Caption = Avanc
    DoEvents
Next

Close filecal
Avancement.Hide

'afficher l'image résultante
Set NewOpen1 = New Resultat
NewOpen1.Image1.Picture = LoadPicture(FichierD3D)
NewOpen1.Caption = FichierD3D

End Sub

