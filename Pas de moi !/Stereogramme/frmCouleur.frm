VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCouleur 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Couleurs du stéréogramme"
   ClientHeight    =   7050
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8160
   Icon            =   "frmCouleur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OuvrPal 
      Caption         =   "Ouvrir palette"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton EnregPal 
      Caption         =   "Enregistrer palette"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6960
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton AideCoul 
      Caption         =   "Aide"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   0
      TabIndex        =   4
      Top             =   6360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1191
      _Version        =   393216
      LargeChange     =   1
      Min             =   2
      Max             =   256
      SelectRange     =   -1  'True
      SelStart        =   2
      TickFrequency   =   10
      Value           =   2
   End
   Begin VB.CommandButton DefCouleur 
      Caption         =   "Définir couleur"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Index           =   0
      Left            =   240
      ScaleHeight     =   19.5
      ScaleMode       =   0  'User
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label NbCoTxt 
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   6360
      Width           =   375
   End
End
Attribute VB_Name = "frmCouleur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NePasFaire As Boolean
Public NoCase As Long
Public SortieValider As Boolean

Private Sub AideCoul_Click()

CD.HelpFile = App.Path & "\aide\" & FichierAide
CD.HelpContext = 5
CD.HelpCommand = cdlHelpContext
CD.ShowHelp

End Sub

Private Sub CancelButton_Click()

SortieValider = False
Unload frmCouleur

End Sub

Private Sub DefCouleur_Click()
Dim R2 As Long
Dim G2 As Long
Dim B2 As Long

MDI3D.CD.ShowColor
    'Cette algorithme transforme la valeur renvoyé par la boite de dialogue,
    'en valeur RGB. en effet, la boite de dialogue renvoit le numéro de la couleur,
    'et nous aurons besoin, par la suite, d'en avoire le code RGB.
    'Il faut d'abord passer par le mode Héxadécimal en convertissant
    'la valeur retourné en valeur héxadécimale, puis, en partant du principe
    'qu'une valeur de couleur Héxadécimale est de type:&HBBGGRR, on peut trouver
    'la valeur R1 représentant la valeur R de la couleur, puis la valeur G1
    'représentant la valeur G de la Couleur, puis la valeur B1 représentant
    'la valeur de G de la couleur. Cela permet de converir ce code en RGB.
    '(attention, il faut inverser car il faut coder en BGR)

R2 = Val("&H" & (Right(Hex(MDI3D.CD.Color), 2)))
If Len(Hex(MDI3D.CD.Color)) >= 4 Then
    G2 = Val("&H" & (Mid(Hex(MDI3D.CD.Color), (Len(Hex(MDI3D.CD.Color))) - 3, 2)))
End If
If Len(Hex(MDI3D.CD.Color)) = 6 Then
    B2 = Val("&H" & Left(Hex(MDI3D.CD.Color), 2))
End If

'sauvegarde dans le tableau
Index = NoCase
PaletteD3D(Index, 1) = B2
PaletteD3D(Index, 2) = G2
PaletteD3D(Index, 3) = R2
    'l'arrière plan de l'image = la couleur renvoyé
Picture1(Index).BackColor = MDI3D.CD.Color
    
    'Interception des erreurs: si l'utilisateur appuie, dans la boite de dialogue,
    'sur Annuler, étant donné que la propriété CancelError de cette boite de
    'dialogue est sur True, la boite de dialogue renvoie une erreur, qui est
    'ici intercépté. si la prpriété CancelError est sur false, si l'utilisateur
    'appuie sur Annuler, la couleur renvoyé par la boite de dialogue est le noire.

Erreur:
If Err.Number = 32755 Then Exit Sub

End Sub

Private Sub EnregPal_Click()
    Dim UdtEnrPalette As FichierPal
    Dim sFile As String
    Dim NumErr As String
    Dim TxtErr As String
    Dim Position As Long

    
    On Error GoTo NouvGestErr1       'active le gestionnaire d'erreurs
    CD.InitDir = App.Path & "\image"  'dossier par défaut
    CD.CancelError = True            'traite le bouton Annuler comme une erreur
    CD.Flags = cdlOFNOverwritePrompt 'affiche un message Remplacer ? si le fichier existe déjà
    CD.Filter = "Palette (*.pal)|*.pal" 'affiche seulement les fichiers se terminant par .Pal
    CD.FileName = " "
    'CD.DialogTitle = TabLng(38)      'titre de la boîte de dialogue
    CD.ShowSave                       'affiche la boite de dialogue Enregistrer sous
    If Len(CD.FileName) = 0 Then
        Exit Sub
    End If
    sFile = CD.FileName


Open sFile For Random As #79 Len = Len(UdtEnrPalette)
'insertion des données dans le fichier
For j = 0 To 15
    For I = 0 To 15
        z = I + 16 * j
        UdtEnrPalette.PalR = PaletteD3D(z, 3)
        UdtEnrPalette.PalG = PaletteD3D(z, 2)
        UdtEnrPalette.PalB = PaletteD3D(z, 1)
        Put #79, z + 1, UdtEnrPalette
    Next
Next
Close #79

Exit Sub                               'empêche le passage au Gestionnaire d'erreurs

NouvGestErr1:
NumErr = Err.Number
TxtErr = Err.Description
'interception de l'erreur "annuler"
If Err.Number <> 32755 Then
    MsgBox ("erreur lors de l'enregistrement: " + NumErr + " " + TxtErr)
End If
Close #79
End Sub

Private Sub Form_Load()
SortieValider = False
NePasFaire = True
Slider1.Value = NbCoulPal
NePasFaire = False
NbCoTxt.Caption = NbCoulPal

'initialisation du language
        frmCouleur.Caption = TabLng(13)
    'boutons
        frmCouleur.OKButton.Caption = TabLng(14)
        frmCouleur.CancelButton.Caption = TabLng(15)
        frmCouleur.AideCoul.Caption = TabLng(17)
        frmCouleur.DefCouleur.Caption = TabLng(19)


'initialisation de l'affichage en fonction du tableau PaletteD3D
For j = 0 To 15
    For I = 0 To 15
        z = I + 16 * j
        a = PaletteD3D(1, 1)
        If z < Slider1.Value Then
            If z Then Load Picture1(z)
            With Picture1(z)
                .Left = Picture1(0).Left + I * 1.5 * Picture1(0).Width
                .Top = Picture1(0).Top + j * 1.5 * Picture1(0).Height
                .BackColor = RGB(PaletteD3D(z, 3), PaletteD3D(z, 2), PaletteD3D(z, 1))
                .Visible = True
                .Enabled = True
            End With
        Else
            If z Then Load Picture1(z)
            With Picture1(z)
                .Left = Picture1(0).Left + I * 1.5 * Picture1(0).Width
                .Top = Picture1(0).Top + j * 1.5 * Picture1(0).Height
                .BackColor = RGB(PaletteD3D(z, 3), PaletteD3D(z, 2), PaletteD3D(z, 1))
                .Visible = False
                .Enabled = False
            End With
        End If
    Next
Next

'sauvegarde du tableau en cas d'annulation de modifs
For j = 1 To 4
    For I = 0 To 255
        SavPaletteD3D(I, j) = PaletteD3D(I, j)
    Next
Next

    ' désactiver feuille principale
    frmOptions.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
If SortieValider = False Then
    'annulation des modifs dans la table palette
    For j = 1 To 4
        For I = 0 To 255
            PaletteD3D(I, j) = SavPaletteD3D(I, j)
        Next
    Next
Else
    'validation du nb de couleurs définies dans la palette
    NbCoulPal = Slider1.Value
End If

     ' activer feuille principale
   frmOptions.Enabled = True

End Sub

Private Sub OKButton_Click()

    SortieValider = True
    
    Unload frmCouleur
End Sub

Private Sub OuvrPal_Click()
    Dim UdtEnrPalette As FichierPal
    Dim sFile As String
    Dim NumErr As String
    Dim TxtErr As String
    Dim Position As Long

    
    On Error GoTo NouvGestErr1       'active le gestionnaire d'erreurs
    CD.InitDir = App.Path & "\image"  'dossier par défaut
    CD.CancelError = True            'traite le bouton Annuler comme une erreur
    CD.Filter = "Palette (*.pal)|*.pal" 'affiche seulement les fichiers se terminant par .Pal
    CD.FileName = " "
    CD.DialogTitle = TabLng(38)      'titre de la boîte de dialogue
    CD.ShowOpen                      'affiche la boite de dialogue Enregistrer sous
    If Len(CD.FileName) = 0 Then
        Exit Sub
    End If
    sFile = CD.FileName


Open sFile For Random As #78 Len = Len(UdtEnrPalette)
'récupération des données du fichier
For j = 0 To 15
    For I = 0 To 15
        z = I + 16 * j
        Get #78, z + 1, UdtEnrPalette
        PaletteD3D(z, 3) = UdtEnrPalette.PalR
        PaletteD3D(z, 2) = UdtEnrPalette.PalG
        PaletteD3D(z, 1) = UdtEnrPalette.PalB
    Next
Next
Close #78

'initialisation de l'affichage en fonction du fichier chargé
For j = 0 To 15
    For I = 0 To 15
        z = I + 16 * j
        With Picture1(z)
            .BackColor = RGB(PaletteD3D(z, 3), PaletteD3D(z, 2), PaletteD3D(z, 1))
        End With
    Next
Next


Exit Sub                               'empêche le passage au Gestionnaire d'erreurs

NouvGestErr1:
NumErr = Err.Number
TxtErr = Err.Description
'interception de l'erreur "annuler"
If Err.Number <> 32755 Then
    MsgBox ("erreur lors de l'enregistrement: " + NumErr + " " + TxtErr)
End If
Close #78

End Sub

Private Sub Picture1_Click(Index As Integer)
'Picture1(NoCase).BorderStyle = 0
NoCase = Index
'Picture1(Index).BorderStyle = 1

End Sub

Private Sub Slider1_Change()

'pour palier à un bug au chargement
If NePasFaire = False Then

NbCoTxt.Caption = Slider1.Value

For j = 0 To 15
    For I = 0 To 15
        z = I + 16 * j
        If z < Slider1.Value Then
            With Picture1(z)
                .Left = Picture1(0).Left + I * 1.5 * Picture1(0).Width
                .Top = Picture1(0).Top + j * 1.5 * Picture1(0).Height
                .BackColor = RGB(PaletteD3D(z, 3), PaletteD3D(z, 2), PaletteD3D(z, 1))
                .Visible = True
                .Enabled = True
            End With
        Else
            With Picture1(z)
                .Left = Picture1(0).Left + I * 1.5 * Picture1(0).Width
                .Top = Picture1(0).Top + j * 1.5 * Picture1(0).Height
                .BackColor = RGB(PaletteD3D(z, 3), PaletteD3D(z, 2), PaletteD3D(z, 1))
                .Visible = False
                .Enabled = False
            End With
        End If
    Next
Next
End If
End Sub
