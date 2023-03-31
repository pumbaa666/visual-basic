VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Pyro-Notes III"
   ClientHeight    =   6255
   ClientLeft      =   1965
   ClientTop       =   1695
   ClientWidth     =   6735
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   600
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":490C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":55E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":62C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6F9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin PN3.CoolBar CoolBar 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   979
      Appearance      =   0
      BackColor       =   14930875
   End
   Begin PN3.Progress Progress 
      Height          =   255
      Left            =   4680
      Top             =   6000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Appearance      =   0
      Value           =   0
      ColorBar        =   15526369
      BackColor       =   14930875
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9340
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Main.frx":7C74
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PN3.Button ButtonRéduire 
      Height          =   195
      Left            =   6000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   30
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      BackColor       =   15921386
      Caption         =   ""
      Pic             =   "Main.frx":7CF0
   End
   Begin PN3.Button ButtonAgrandir 
      Height          =   195
      Left            =   6240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      BackColor       =   15921386
      Caption         =   ""
      Pic             =   "Main.frx":89CA
   End
   Begin PN3.Button ButtonQuitter 
      Height          =   195
      Left            =   6480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   30
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      BackColor       =   15921386
      Caption         =   ""
      Pic             =   "Main.frx":96A4
   End
   Begin VB.Image ImageResize 
      Height          =   240
      Left            =   6480
      Picture         =   "Main.frx":A37E
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   240
   End
   Begin VB.Shape ShapeResize 
      BackColor       =   &H00E3D3BB&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   6480
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label LabelTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "  Pyro-Notes III"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
   End
   Begin VB.Shape ShapeTitle 
      BackColor       =   &H007D631C&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3D3BB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nouveau texte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   6735
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'############################################################################
'#Infos :               Nom : Pyro-Notes III                                #
'#=======               Programmeur : PyroSmoke                             #
'#                      E-Mail : pyrosmoke@hotmail.com                      #
'#                      IRC : #pyroworld@irc.espry.org                      #
'############################################################################
'#Reste à faire :       [ ] Taille                                          #
'#===============       [ ] Gras                                            #
'#                      [ ] Italique                                        #
'#                      [ ] Souligné                                        #
'#                      [ ] Barré                                           #
'#                      [ ] Edition différente pour txt et rtf              #
'#                      [/] Contrôle Menu                                   #
'#                      [/] Prise en charge des polices                     #
'#                      [ ] Highlight pour la CoolBar                       #
'#                      [ ] Coolbar : AddItem pendant la programmation      #
'#                      [ ] Régler bug d'affichage pour le resize           #
'#                      [ ] Faire système de mise à jour                    #
'#                      [ ] Mettre FlatScrollBar pour la RitchTextBox       #
'#                      [ ] Réfléchir à un système de scripting             #
'#                      [/] Fonction d'impression                           #
'#                      [ ] Sauvegarde temporaire du fichier                #
'#                      [ ] Mettre la textbox translucide                   #
'#                      [/] Revoir le procédé de cryptage                   #
'#                      [ ] Activer les messages d'erreurs sur tout PN3     #
'#                      [ ] Possibilité de conversion Win<=>Unix            #
'#                      [ ] Bug d'affichage pour la message box             #
'#                      [ ] Revoir les assignations de fichiers             #
'#                      [ ] Activer le LastDir pour le Common               #
'#                      [ ] Plantage quand resize trop petit                #
'############################################################################
'#Optimisation :        [ ] Utiliser le moins de variables possible         #
'#==============        [ ] Ne pas mettre, si possible, les variables en    #
'#                          mode public                                     #
'#                      [ ] Convertir les bmp stockés en jpg                #
'#                      [ ] Virer les lignes de code inutile (répétitions)  #
'#                      [ ] Virer les contrôles qui ne servent à rien       #
'#                      [ ] Laisser dans les modules les fonctions qui      #
'#                          servent pour plusieurs Feuilles                 #
'#                      [ ] Vérifier les variables (Private, Public ...)    #
'#                      [ ] Passer tous les contrôles en OCX                #
'#                      [ ] Tout les var reg avec une valeur par défaut     #
'#                      [ ] Une seule fonction si plusieurs semblables      #
'#                      [ ] Faire gaffe au BufferOverflow                   #
'############################################################################

'Variables générales
Dim LastLeft, LastTop, LastWidth, LastHeight As Single
Dim LastDir As String
Dim ResizeIt As Boolean
Dim XBase, YBase, NewX, NewY As Single

Private Sub ButtonAgrandir_Click()

If Me.WindowState = 0 Then
    'Sauvegarde de la grandeur de base pour la restauration de la fenêtre
    LastLeft = Me.Left
    LastTop = Me.Top
    LastWidth = Me.Width
    LastHeight = Me.Height
    Me.WindowState = 2
Else
    Me.WindowState = 0
    Me.Move LastLeft, LastTop, LastWidth, LastHeight
End If

End Sub

Private Sub ButtonQuitter_Click()

Quit

End Sub

Private Sub ButtonRéduire_Click()

Me.WindowState = 1

End Sub

Private Sub CoolBar_ButtonClick(Key As String)

'Appui sur un bouton de la CoolBar

Select Case Key
Case "New"
    NewFile
Case "Open"
    OpenFile
Case "Save"
    SaveFile
Case "Search"
    Me.Enabled = False
    Search.Show 1
Case "Print"
    'Printer.Print ; Text.Text
    'Printer.EndDoc
Case "Crypt"
    Me.Enabled = False
    Crypt.Show 1
Case "Options"
    Me.Enabled = False
    Config.Show 1
Case "About"
    About
Case "Quit"
    Quit
End Select

End Sub

Private Sub Form_Load()

'Redéfinition du titre avec le numéro actuel de version
NumVersion = "- Beta " & App.Major & "." & App.Minor & "." & App.Revision
LabelTitle.Caption = "  Pyro-Notes III " & NumVersion

'En cas de premier démarrage : assignation de fichiers et enregistrement de paramètres
If GetSetting("Pyro-Notes III", "Config", "FirstStart") <> "Yes" Then
    If MessageBox.Message("Ceci est le premier démarrage de Pyro-Notes III." & vbCrLf & "Voulez-vous que l'ouverture des fichiers textes soit pris en charge par Pyro-Notes III?", "Premier lancement", YesNo, Request, Main) = Yes Then AssignPN3TXT: AssignPN3RTF
    SaveBaseParams
End If

'Sauvegarde de paramètres utiles au redimensionnage
Longueur = Me.ScaleWidth - (ButtonQuitter.Left + ButtonQuitter.Width)

'Chargement de la liste des polices
'LoadPolices

'Chargement des paramètres si PN3 a déjà démarré ou alors enregistrement
If GetSetting("Pyro-Notes III", "Config", "FirstStart") = "Yes" Then
    LoadParams
Else
    SaveSetting "Pyro-Notes III", "Config", "FirstStart", "Yes"
End If

'Chargement de la CoolBar
CoolBar.AddButton 1, ImageList, "New", "Nouveau"
CoolBar.AddButton 2, ImageList, "Open", "Ouvrir"
CoolBar.AddButton 3, ImageList, "Save", "Sauver"
CoolBar.AddButton 4, ImageList, "Search", "Rechercher"
CoolBar.AddButton 5, ImageList, "Print", "Imprimer"
CoolBar.AddButton 6, ImageList, "Crypt", "Cryptage"
CoolBar.AddButton 7, ImageList, "Options", "Options"
CoolBar.AddButton 8, ImageList, "About", "A propos..."
CoolBar.AddButton 9, ImageList, "Quit", "Fermer"

Me.Show
DoEvents

'Regarde si le programme a été appelé par ouverture d'un fichier
If Command() <> "" Then
    Status "Ouverture en cours..."
    FichierActuel = Command()
    LoadFileZ FichierActuel
End If

Me.Show
Text.SetFocus

End Sub

Private Sub Form_Resize()

'Ne pas redimensionner en réduction
If Me.WindowState <> 1 Then
    'Pour éviter que le prog plante si on le redimmensionne trop petit
    If Me.Width < 3000 Then Me.Width = 3000
    If Me.Height < 3000 Then Me.Height = 3000
    'On redimmensionne correctement les contrôles
    ReMap
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Quit

End Sub

Private Sub ImageResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Sauvegarde la position de base de la souris
If Button = vbKeyLButton Then
    XBase = ImageResize.Width - X
    YBase = ImageResize.Height - Y
End If

End Sub

Private Sub ImageResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'On sauvegarde la nouvelle dimension de la fenêtre
If Button = vbKeyLButton Then
    Me.Width = ImageResize.Left + X + XBase
    Me.Height = ImageResize.Top + Y + YBase
End If

End Sub

Private Sub LabelTitle_DblClick()

ButtonAgrandir_Click

End Sub

Private Sub Text_Change()

'Affichage de la grosseur du texte
If Len(Text.Text) < 1024 Then LabelTitle.Caption = "  Pyro-Notes III " & NumVersion & " - " & Len(Text.Text) & " octets"
If Len(Text.Text) >= 1024 Then LabelTitle.Caption = "  Pyro-Notes III " & NumVersion & " - " & Round(Len(Text.Text) / 1024, 1) & " Ko"

End Sub

Private Sub Text_KeyDown(KeyCode As Integer, Shift As Integer)

'Recherche suivante
If KeyCode = vbKeyF3 Then
    If Search.SearchNextText = False Then MessageBox.Message "La recherche n'a rien donné.", "Recherche infructueuse", OkOnly, Information, Main: Text.SetFocus
End If

End Sub

Private Sub Text_KeyUp(KeyCode As Integer, Shift As Integer)

'Changement de valeur de la variable si on change le texte
If TamponTexte <> Text.Text And MustBeSaved = False Then MustBeSaved = True

End Sub

Private Sub Text_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour afficher le menu
'If Button = vbKeyRButton Then Me.PopupMenu Menu, , X, Y

End Sub

Private Sub NewFile()

'Changements pour un nouveau texte
Text.Text = ""
Status "Nouveau texte"
AlreadySaved = False
FichierActuel = ""
MustBeSaved = False
Text.SetFocus

End Sub

Private Sub OpenFile()

Dim Temp As String

'Ouverture d'un fichier
If LastDir <> "" Then
    Temp = Common.ShowCommon("Ouvrir", LastDir)
Else
    Temp = Common.ShowCommon("Ouvrir")
End If
If Temp <> "" Then
    LastDir = Temp
    FichierActuel = Temp
    LoadFileZ Temp
End If
Text.SetFocus

End Sub

Private Sub SaveFile()

'Sauvegarde rapide
If AlreadySaved = False Then SaveFileAs: Exit Sub
SaveFileZ Text.Text, FichierActuel
Text.SetFocus

End Sub

Private Sub SaveFileAs()

Dim Temp As String

'Sauvegarde

If LastDir <> "" Then
    Temp = Common.ShowCommon("Enregistrer sous", LastDir)
Else
    Temp = Common.ShowCommon("Enregistrer sous")
End If
If Temp <> "" Then
    LastDir = Temp
    FichierActuel = Temp
    SaveFileZ Text.Text, Temp
End If
Text.SetFocus

End Sub

Private Sub About()

'MessageBox pour les renseignements sur le programmeur
MessageBox.Message "Programmé par PyroSmoke" & vbCrLf & vbCrLf & _
"Pour tout renseignement ou rapport de bug : pyrosmoke@hotmail.com." & vbCrLf & _
vbCrLf & "IRC : #pyroworld@irc.espry.org", "A propos de ce somptueux logiciel", OkOnly, Information, Main

Text.SetFocus

End Sub

Private Sub Quit()

'Vérification du besoin d'enregistrer le travail avant de quitter
If MustBeSaved = True Then
    If FichierActuel = "" Then
        Select Case MessageBox.Message("Votre travail a été modifié." & vbCrLf & "Voulez-vous sauvegarder les dernières modifications?", "Fichier non sauvegardé", YesNocancel, Request, Main)
        Case Yes
            SaveFile
        Case Cancel
            Exit Sub
        End Select
    Else
        Select Case MessageBox.Message(FichierActuel & " a été modifié." & vbCrLf & "Voulez-vous sauvegarder les dernières modifications?", "Fichier non sauvegardé", YesNocancel, Request, Main)
        Case Yes
            SaveFile
        Case Cancel
            Exit Sub
        End Select
    End If
End If

'Sauvegarde de l'emplacement de la fenêtre
SaveSetting "Pyro-Notes III", "Config", "Left", Me.Left
SaveSetting "Pyro-Notes III", "Config", "Top", Me.Top

End

End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour bouger la fenêtre
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&

End Sub
