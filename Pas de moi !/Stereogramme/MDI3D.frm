VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDI3D 
   BackColor       =   &H8000000A&
   Caption         =   "Dessin3D"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9210
   Icon            =   "MDI3D.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   8520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Fichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Ouvrir 
         Caption         =   "&Ouvrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu Quitter 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Creer3D 
      Caption         =   "&Stéréogramme"
      Begin VB.Menu CD3D 
         Caption         =   "&Créer"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Options 
         Caption         =   "&Options..."
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu Fenetre 
      Caption         =   "F&enêtre"
      WindowList      =   -1  'True
   End
   Begin VB.Menu Aide 
      Caption         =   "&?"
      Begin VB.Menu AidePgm 
         Caption         =   "&Aide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Aproposde 
         Caption         =   "À &propos de..."
      End
   End
   Begin VB.Menu menucache2 
      Caption         =   "me&nu caché2"
      Visible         =   0   'False
      Begin VB.Menu propriete 
         Caption         =   "&propriétés"
      End
   End
End
Attribute VB_Name = "MDI3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewOpen As Form
Dim NewOpen1 As Form

Private Sub AidePgm_Click()
Dim Aide As Long
Dim Chemin As String

Chemin = "c:\windows\winhlp32.exe " + App.Path + "\aide\" & FichierAide
Aide = Shell(Chemin, vbNormalFocus)
End Sub

Private Sub Aproposde_Click()
frmAbout.Show
End Sub

Private Sub CD3D_Click()
Dim UdtEnrFichier As EnteteFichierBmp
Dim Msg As String

FichierNomLongAide = App.Path & "\aide\" & FichierAide
'----------------------------------------------------------------'
'    ***  Vérifie la validité du fichier à traiter         ***   '
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
    Get #1, 1, UdtEnrFichier
    Close #1
    
    'si le fichier est non bitmap, sortie sans création du stéréogramme
    If UdtEnrFichier.EFBFileType <> "BM" Then
        Msg = "Fichier non Bitmap."
        a = MsgBox(Msg, vbMsgBoxHelpButton, "Erreur!", FichierNomLongAide, 6)
        Exit Sub
    End If
    'si le fichier est bitmap mais non 8 Bits, sortie sans création du stéréogramme
    If UdtEnrFichier.EFBBitsPerPixel <> "8" Then
        Msg = "Utilisez un fichier Bitmap 8 bits (niveaux de gris)." & Chr(13) & _
              "Le fichier actuel est en mode " & UdtEnrFichier.EFBBitsPerPixel & " Bits."
        a = MsgBox(Msg, vbMsgBoxHelpButton, "Erreur!", FichierNomLongAide, 6)
        Exit Sub
    End If
    'si le fichier est un bitmap compressé, sortie sans création du stéréogramme
    If UdtEnrFichier.EFBCompression <> "0" Then
        Msg = "Utilisez un fichier Bitmap non compressé."
        a = MsgBox(Msg, vbMsgBoxHelpButton, "Erreur!", FichierNomLongAide, 6)
        Exit Sub
    End If
    
'----------------------------------------------------------------'
'    ***  Création du Bitmap Stéréogramme                  ***   '
'----------------------------------------------------------------'

CreationBmpNivGris

End Sub

Private Sub MDIForm_Load()
LrgBnd = 30
NbProf = 64
FctNbProf = 4
NbCoulPal = 2
PaletteD3D(0, 1) = 255
PaletteD3D(0, 2) = 255
PaletteD3D(0, 3) = 255
PaletteD3D(0, 4) = 0
PaletteD3D(1, 1) = 0
PaletteD3D(1, 2) = 0
PaletteD3D(1, 3) = 0
PaletteD3D(1, 4) = 0


'récupération des réglages
    I = 0
    Fic = App.Path & "\reglages.ini"
    Open Fic For Input As #98
    Do While Not EOF(98)
        Line Input #98, L
        If UCase(Left(L, 4)) = "LNG=" Then
            Lng = Mid(L, 5)
        End If
        If UCase(Left(L, 7)) = "LRGBND=" Then
            LrgBnd = Mid(L, 8)
        End If
        If UCase(Left(L, 10)) = "FCTNBPROF=" Then
            FctNbProf = Mid(L, 11)
        End If
        If UCase(Left(L, 10)) = "NBCOULPAL=" Then
            NbCoulPal = Mid(L, 11)
        End If
        If UCase(Left(L, 7)) = "NBPROF=" Then
            NbProf = Mid(L, 8)
        End If
        If UCase(Left(L, 11)) = "RESULTAT3D=" Then
            Resultat3D = Mid(L, 12)
        End If
        If UCase(Left(L, 8)) = "TEXTURE=" Then
            FichierNomLongTexture = Mid(L, 9)
        End If
    Loop
    Close #98
    'maj des données texture
    If FichierNomLongTexture <> "" Then
        MajTexture
    End If
    'ouverture du fichier langue
    If Lng = "" Then
        Lng = "Français"
    End If
        Fic = App.Path & "\langue\" & Lng
        Open Fic For Input As #99
        'remplissage du tableau langues
        For I = 1 To 40
            Line Input #99, L
            TabLng(I) = L
        Next
        Close #99
        'entêtes
        MDI3D.Caption = TabLng(1)
        
        'menus
        MDI3D.Fichier.Caption = TabLng(2)
        MDI3D.Ouvrir.Caption = TabLng(3)
        MDI3D.Quitter.Caption = TabLng(4)
        MDI3D.Creer3D.Caption = TabLng(5)
        MDI3D.CD3D.Caption = TabLng(6)
        MDI3D.Options.Caption = TabLng(7)
        MDI3D.Fenetre.Caption = TabLng(8)
        MDI3D.Aide.Caption = TabLng(9)
        MDI3D.AidePgm.Caption = TabLng(10)
        MDI3D.Aproposde.Caption = TabLng(11)

        FichierAide = TabLng(39)

    MDI3D.CD3D.Enabled = False

End Sub

Private Sub Options_Click()
frmOptions.Show
End Sub

Private Sub Ouvrir_Click()
    Dim UdtEnrFichier As EnteteFichierBmp
    Dim UdtEnrCorpsFichier As CorpsFichierBmp
    Dim sFile As String
    Dim NumErr As String
    Dim TxtErr As String
    Dim Position As Long

'ajouter un test pour savoir si on veut écraser un fichier ouvert
'par un autre

    On Error GoTo NouvGestErr1       'active le gestionnaire d'erreurs
    CD.InitDir = App.Path & "\image"  'dossier par défaut
    CD.CancelError = True            'traite le bouton Annuler comme une erreur
    CD.Filter = "Data (*.BMP)|*.BMP" 'affiche seulement les fichiers se terminant par .BMP
    CD.FileName = " "
    CD.DialogTitle = TabLng(38)      'titre de la boîte de dialogue
    CD.ShowOpen                      'affiche la boite de dialogue Enregistrer sous
    If Len(CD.FileName) = 0 Then
        Exit Sub
    End If
    sFile = CD.FileName


Set NewOpen1 = New Resultat
NewOpen1.Image1.Picture = LoadPicture(sFile)
NewOpen1.Caption = sFile

Exit Sub
NouvGestErr1:
NumErr = Err.Number
TxtErr = Err.Description
'interception de l'erreur "annuler"
If Err.Number <> 32755 Then
    MsgBox ("erreur lors de l'ouverture: " + NumErr + " " + TxtErr)
End If
End Sub

Private Sub Quitter_Click()
End
End Sub

Private Sub Fermer_Click()
Unload Resultat
End Sub

Private Sub propriete_click()
    Dim UdtEnrFichier As EnteteFichierBmp
    
    Open FormActive For Random As #2 Len = Len(UdtEnrFichier)
    Get #2, 1, UdtEnrFichier

    Set NewOpen = New EnteteF

    NewOpen.Caption = FormActive
    NewOpen.Text1 = UdtEnrFichier.EFBFileType
    NewOpen.Text2 = UdtEnrFichier.EFBFileSize
    NewOpen.Text3 = UdtEnrFichier.EFBReserved
    NewOpen.Text4 = UdtEnrFichier.EFBBitMapOffset
    NewOpen.Text5 = UdtEnrFichier.EFBHeaderSize
    NewOpen.Text6 = UdtEnrFichier.EFBWidth
    NewOpen.Text7 = UdtEnrFichier.EFBHeight
    NewOpen.Text8 = UdtEnrFichier.EFBPlanes
    NewOpen.Text9 = UdtEnrFichier.EFBBitsPerPixel
    NewOpen.Text10 = UdtEnrFichier.EFBCompression
    NewOpen.Text11 = UdtEnrFichier.EFBSizeOfBitMap
    NewOpen.Text12 = UdtEnrFichier.EFBHorzResolution
    NewOpen.Text13 = UdtEnrFichier.EFBVertResolution
    NewOpen.Text14 = UdtEnrFichier.EFBColorsUsed
    NewOpen.Text15 = UdtEnrFichier.EFBColorsImportant

Close #2
End Sub
