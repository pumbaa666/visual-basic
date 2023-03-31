VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5385
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6405
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Tram 
      Caption         =   "Trame de fond"
      Height          =   735
      Left            =   360
      TabIndex        =   27
      Top             =   2280
      Width           =   5655
      Begin VB.OptionButton PAlTram 
         Caption         =   "Trame"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton PAlTram 
         Caption         =   "Palette"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Couleurs 
         Caption         =   "Couleurs"
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Définition des couleurs du stéréogramme"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ListBox Langue 
      Height          =   840
      Left            =   4920
      TabIndex        =   25
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dessin"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   2640
      TabIndex        =   21
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "Redimensionne le cadre"
         Enabled         =   0   'False
         Height          =   435
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Redimensionne le dessin"
         Enabled         =   0   'False
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Libre"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Le dessin et le cadre sont indépendants"
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton AideOpt 
      Caption         =   "Aide"
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ecart des yeux"
      Height          =   1455
      Left            =   360
      TabIndex        =   13
      Top             =   3120
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   19
         ToolTipText     =   "-5"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   18
         ToolTipText     =   "+5"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   17
         ToolTipText     =   "+1"
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   16
         ToolTipText     =   "-1"
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtLargeur 
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   80
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Niveaux de profondeur"
      Height          =   1935
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Ultra!"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   26
         ToolTipText     =   "128 niv. de profondeur"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Très fin"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "128 niv. de profondeur"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fin"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         ToolTipText     =   "64 niv. de profondeur"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Standard"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   "32 niv. de profondeur"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Exemple 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Exemple 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Exemple 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AideOpt_Click()

CD.HelpFile = App.Path & "\aide\" & FichierAide
CD.HelpContext = 2
CD.HelpCommand = cdlHelpContext
CD.ShowHelp

End Sub

Private Sub cmdApply_Click()

    If Option1(1).Value = True Then NbProf = 32
    If Option1(2).Value = True Then NbProf = 64
    If Option1(3).Value = True Then NbProf = 128
    If Option1(4).Value = True Then NbProf = 256
    If Option1(1).Value = True Then FctNbProf = 8
    If Option1(2).Value = True Then FctNbProf = 4
    If Option1(3).Value = True Then FctNbProf = 2
    If Option1(4).Value = True Then FctNbProf = 1
    LrgBnd = Slider1.Value
    
    If PAlTram(0).Value = True Then
        Resultat3D = "Random"
    Else
        Resultat3D = "Texture"
    End If

    cmdApply.Enabled = False
    
    'MAJ des réglages dans le fichier réglages.ini (module SubRoutines)
    MajRegIni

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Option1(1).Value = True Then NbProf = 32
    If Option1(2).Value = True Then NbProf = 64
    If Option1(3).Value = True Then NbProf = 128
    If Option1(4).Value = True Then NbProf = 256
    If Option1(1).Value = True Then FctNbProf = 8
    If Option1(2).Value = True Then FctNbProf = 4
    If Option1(3).Value = True Then FctNbProf = 2
    If Option1(4).Value = True Then FctNbProf = 1
    LrgBnd = Slider1.Value
    
    If PAlTram(0).Value = True Then
        Resultat3D = "Random"
    Else
        Resultat3D = "Texture"
    End If

   
    'MAJ des réglages dans le fichier réglages.ini (module SubRoutines)
    MajRegIni

    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Temp As Integer

Temp = txtLargeur.Text

If Index = 1 Then
    If Temp < 60 Then
        Temp = Temp + 1
    End If
Else
    If Index = 0 Then
        If Temp > 1 Then
            Temp = Temp - 1
        End If
    Else
        If Index = 3 Then
            If Temp < 56 Then
                Temp = Temp + 5
            End If
        Else
            If Temp > 5 Then
                Temp = Temp - 5
            End If
        End If
    End If
End If
txtLargeur.Text = Temp
Slider1.Value = Temp
cmdApply.Enabled = True

End Sub

Private Sub Couleurs_Click()
If frmOptions.Couleurs.Caption = TabLng(18) Then
    frmCouleur.Show
End If
If frmOptions.Couleurs.Caption = TabLng(40) Then
    frmTexture.Show
End If
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim Fic As String
Dim Dossier As String

    'met la bonne langue
        'entête
        frmOptions.Caption = TabLng(12)
        
        'boutons
        frmOptions.cmdOK.Caption = TabLng(14)
        frmOptions.cmdCancel.Caption = TabLng(15)
        frmOptions.cmdApply.Caption = TabLng(16)
        frmOptions.AideOpt.Caption = TabLng(17)
        'doit on afficher couleurs (random) ou Images (image)
        If Resultat3D = "Random" Then
            frmOptions.Couleurs.Caption = TabLng(18)
            PAlTram(0).Value = True
        Else
            frmOptions.Couleurs.Caption = TabLng(40)
            PAlTram(1).Value = True
        End If
        
        'autres textes
        frmOptions.Frame1.Caption = TabLng(20)
        frmOptions.Frame2.Caption = TabLng(22)
        frmOptions.Frame3.Caption = TabLng(21)
        frmOptions.Option1(1).Caption = TabLng(23)
        frmOptions.Option1(1).ToolTipText = TabLng(24)
        frmOptions.Option1(2).Caption = TabLng(25)
        frmOptions.Option1(2).ToolTipText = TabLng(26)
        frmOptions.Option1(3).Caption = TabLng(27)
        frmOptions.Option1(3).ToolTipText = TabLng(28)
        frmOptions.Option2(0).Caption = TabLng(29)
        frmOptions.Option2(1).Caption = TabLng(30)
        frmOptions.Option2(2).Caption = TabLng(31)
        frmOptions.Command1(0).ToolTipText = TabLng(33)
        frmOptions.Command1(1).ToolTipText = TabLng(34)
        frmOptions.Command1(2).ToolTipText = TabLng(32)
        frmOptions.Command1(3).ToolTipText = TabLng(35)
        frmOptions.Couleurs.ToolTipText = TabLng(36)
        frmOptions.Option2(0).ToolTipText = TabLng(37)
    
    'récupère la liste des fichiers langue
    Dossier = App.Path & "\langue\*"
    Fic = Dir(Dossier)
    Langue.AddItem Fic
    While Fic <> ""
        Fic = Dir
        Langue.AddItem Fic
    Wend
    Langue.Text = Lng
    
    ' désactiver feuille principale
    MDI3D.Enabled = False
    
    ' Cntre la feuille.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    'initialisation des variables
    If NbProf = 32 Then
        Option1(1).Value = True
    End If
    If NbProf = 64 Then
        Option1(2).Value = True
    End If
    If NbProf = 128 Then
        Option1(3).Value = True
    End If
    If NbProf = 256 Then
        Option1(4).Value = True
    End If
    
    Slider1.Value = LrgBnd
    txtLargeur.Text = LrgBnd
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' activer feuille principale
    MDI3D.Enabled = True

End Sub

Private Sub Langue_Click()
Dim Fic As String
Dim L As String
    
    'MAJ des réglages de la langue
    MajLngIni
    
        'entêtes
        MDI3D.Caption = TabLng(1)
        frmOptions.Caption = TabLng(12)
        
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
        
        'boutons
        frmOptions.cmdOK.Caption = TabLng(14)
        frmOptions.cmdCancel.Caption = TabLng(15)
        frmOptions.cmdApply.Caption = TabLng(16)
        frmOptions.AideOpt.Caption = TabLng(17)
        'doit on afficher couleurs (random) ou Trames (image)
        If PAlTram(0).Value = True Then
            frmOptions.Couleurs.Caption = TabLng(18)
        Else
            frmOptions.Couleurs.Caption = TabLng(40)
        End If
        
        frmOptions.Frame1.Caption = TabLng(20)
        frmOptions.Frame2.Caption = TabLng(22)
        frmOptions.Frame3.Caption = TabLng(21)
        frmOptions.Option1(1).Caption = TabLng(23)
        frmOptions.Option1(1).ToolTipText = TabLng(24)
        frmOptions.Option1(2).Caption = TabLng(25)
        frmOptions.Option1(2).ToolTipText = TabLng(26)
        frmOptions.Option1(3).Caption = TabLng(27)
        frmOptions.Option1(3).ToolTipText = TabLng(28)
        frmOptions.Option2(0).Caption = TabLng(29)
        frmOptions.Option2(1).Caption = TabLng(30)
        frmOptions.Option2(2).Caption = TabLng(31)
        frmOptions.Command1(0).ToolTipText = TabLng(33)
        frmOptions.Command1(1).ToolTipText = TabLng(34)
        frmOptions.Command1(2).ToolTipText = TabLng(32)
        frmOptions.Command1(3).ToolTipText = TabLng(35)
        frmOptions.Couleurs.ToolTipText = TabLng(36)
        frmOptions.Option2(0).ToolTipText = TabLng(37)
        
        FichierAide = TabLng(39)
    
End Sub

Private Sub Option1_Click(Index As Integer)
cmdApply.Enabled = True
End Sub

Private Sub Option2_Click(Index As Integer)
cmdApply.Enabled = True
End Sub

Private Sub PAlTram_Click(Index As Integer)
cmdApply.Enabled = True
If PAlTram(0).Value = True Then
    Couleurs.Caption = TabLng(18)
Else
    Couleurs.Caption = TabLng(40)
End If
End Sub

Private Sub Slider1_Change()
cmdApply.Enabled = True
txtLargeur.Text = Slider1.Value

End Sub
