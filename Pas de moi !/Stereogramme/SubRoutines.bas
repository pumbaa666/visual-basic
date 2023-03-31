Attribute VB_Name = "SubRoutines"
Sub MajRegIni()
    'MAJ des réglages dans le fichier réglages.ini
    I = 0
    Fic = App.Path & "\reglages.ini"
    Open Fic For Input As #98
    'récupération des infos
    Do While Not EOF(98)
        Line Input #98, L
        If UCase(Left(L, 4)) = "LNG=" Then
            If frmOptions.Langue.Text = "" Then
                L = "Lng=" & Lng
            Else
                L = "Lng=" & frmOptions.Langue.Text
            End If
        End If
        If UCase(Left(L, 7)) = "NBPROF=" Then
            L = "NbProf=" & NbProf
        End If
        If UCase(Left(L, 10)) = "FCTNBPROF=" Then
            L = "FctNbProf=" & FctNbProf
        End If
        If UCase(Left(L, 7)) = "LRGBND=" Then
            L = "LrgBnd=" & LrgBnd
        End If
        If UCase(Left(L, 11)) = "RESULTAT3D=" Then
            L = "Resultat3D=" & Resultat3D
        End If
        If UCase(Left(L, 8)) = "TEXTURE=" Then
            L = "TEXTURE=" & FichierNomLongTexture
        End If
        
        I = I + 1
        TabRg(I) = L
    Loop
    Close #98
    Imax = I
    'MAJ proprement dite
    Open Fic For Output As #98
    For I = 1 To Imax
        L = TabRg(I)
        Print #98, L
    Next
    Close #98
End Sub

Sub MajLngIni()
    'ouverture du fichier langue
    Fic = App.Path & "\langue\" & frmOptions.Langue.Text
    Open Fic For Input As #99
    'remplissage du tableau langues
    For I = 1 To 39
        Line Input #99, L
        TabLng(I) = L
    Next
    Close #99

    'MAJ des réglages de la langue dans le .ini
    I = 0
    Fic = App.Path & "\reglages.ini"
    Open Fic For Input As #98
    Do While Not EOF(98)
        Line Input #98, L
        If UCase(Left(L, 4)) = "LNG=" Then
            L = "Lng=" & frmOptions.Langue.Text
            Lng = frmOptions.Langue.Text
        End If
        I = I + 1
        TabRg(I) = L
    Loop
    Close #98
    Imax = I
    Open Fic For Output As #98
    For I = 1 To Imax
        L = TabRg(I)
        Print #98, L
    Next
    Close #98
End Sub

Sub MajTexture()
Dim UdtEnrFichier As EnteteFichierBmp
Dim UdtEnrCorpsFichier As CorpsFichierBmp
Dim Position As Long

'récupère la taille dans l'entête
Open FichierNomLongTexture For Random As #73 Len = Len(UdtEnrFichier)
    Get #73, 1, UdtEnrFichier
Close #73
TexHResol = UdtEnrFichier.EFBWidth
TexVResol = UdtEnrFichier.EFBHeight
 
'récupération de la palette de couleur
Open FichierNomLongTexture _
    For Random As #73 Len = Len(UdtEnrCorpsFichier)
    Indice = 54
    For I = 0 To 255    '255
        For J = 1 To 4
            Indice = Indice + 1
            Get #73, Indice, UdtEnrCorpsFichier
            PaletteTexD3D(I, J) = UdtEnrCorpsFichier.CouleurPixel
        Next
    Next

    'récupère l'image proprement dite
    ReDim TexCorps(TexVResol, TexHResol)
    Position = UdtEnrFichier.EFBBitMapOffset
    For CY = TexVResol To 1 Step -1
        For CX = 1 To TexHResol
            Position = Position + 1
            Get #73, Position, UdtEnrCorpsFichier
            TexCorps(CY, CX) = UdtEnrCorpsFichier.CouleurPixel
        Next
    Next
Close #73

End Sub
