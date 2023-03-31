VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surveillance video"
   ClientHeight    =   5010
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5355
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dial1 
      Left            =   240
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer ClkDelay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   4320
   End
   Begin VB.CommandButton CmdDegre 
      Caption         =   "-50°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton CmdDegre 
      Caption         =   "-10°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4200
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton CmdDegre 
      Caption         =   "-5°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CmdDegre 
      Caption         =   "+50°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton CmdDegre 
      Caption         =   "+10°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton CmdDegre 
      Caption         =   "+5°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Valider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtDegre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer ClkEnvoie 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   4320
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      NullDiscard     =   -1  'True
      ParityReplace   =   48
   End
   Begin VB.Timer ClkCam 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   3600
   End
   Begin VB.Label Label2 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Tourner la caméra à"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Tourner la caméra de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Shape ShpCam 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   1800
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label LblDegre 
      Caption         =   "0°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape ShpRay 
      Height          =   2175
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label LblTitre 
      Alignment       =   2  'Center
      Caption         =   "Programme de surveillance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu MenuNew 
         Caption         =   "Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu Tiret1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuVoir 
         Caption         =   "Voir l'image"
         Shortcut        =   ^I
      End
      Begin VB.Menu Tiret2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuQuit 
         Caption         =   "Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuOutils 
      Caption         =   "&Outils"
      Begin VB.Menu MenuCapture 
         Caption         =   "Capture d'image"
      End
      Begin VB.Menu MenuDetail 
         Caption         =   "Détails d'image"
      End
      Begin VB.Menu Tiret3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPas 
         Caption         =   "Nb Pas du moteur"
      End
      Begin VB.Menu MenuVit 
         Caption         =   "Vitesse"
      End
      Begin VB.Menu Tiret4 
         Caption         =   "-"
      End
      Begin VB.Menu MenuVisser 
         Caption         =   "Visser la WebCam"
         Shortcut        =   ^V
      End
      Begin VB.Menu MenuDevisser 
         Caption         =   "Dévisser la WebCam"
         Shortcut        =   ^D
      End
      Begin VB.Menu Tiret5 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCouleur 
         Caption         =   "Couleur de la WebCam"
      End
   End
   Begin VB.Menu MenuAide 
      Caption         =   "&Aide"
      Begin VB.Menu MenuAidePrinc 
         Caption         =   "?"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Tiret6 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPropos 
         Caption         =   "A propos..."
      End
   End
   Begin VB.Menu MenuPopup 
      Caption         =   "&Qu'est-ce que c'est ?!?"
      Visible         =   0   'False
      Begin VB.Menu MenuPopupAide 
         Caption         =   "Qu'est-ce que c'est ?"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vValTot As Integer
Dim vNbVir As Boolean
Dim vDegre As Double
Dim vDestination As Double
Dim vDegText As Single
Dim vFonction As Integer
Dim vPopup As String

Private Sub ClkCam_Timer()
Dim vCount As Integer
Dim vRayon As Integer
Dim vXRay As Integer
Dim vYRay As Integer
Dim vTemp As Boolean

vRayon = ShpRay.Width / 2
vXRay = ShpRay.Left + ShpRay.Width / 2
vYRay = (ShpRay.Top + ShpRay.Height / 2) - 100
    
' Désactive les boutons pour qu'on ne puisse pas cliquer plusieurs fois de suite
    CmdOk.Enabled = False
    For vCount = 0 To 5
        CmdDegre(vCount).Enabled = False
    Next
    
' Arrete le moteur si il a à peu près atteint sa destination
    If vDegre - vDestination > -0.01 And vDegre - vDestination < 0.01 Then
        ClkCam.Enabled = False
        CmdOk.Enabled = True
        For vCount = 0 To 5
            CmdDegre(vCount).Enabled = True
        Next
        LblDegre.Caption = vDegText & "°"
        
' Enregistre l'image si le moteur atteint une des trois positions
        If vValTot = tArret(0) Then
            frmMainCam.VideoPortal1.PictureToFile 0, 24, "c:\temp\image1.bmp", ""
        ElseIf vValTot = tArret(1) Then
            frmMainCam.VideoPortal1.PictureToFile 0, 24, "c:\temp\image2.bmp", ""
        ElseIf vValTot = tArret(2) Then
            frmMainCam.VideoPortal1.PictureToFile 0, 24, "c:\temp\image3.bmp", ""
        Else
            ClkDelay.Enabled = False
        End If
        Exit Sub
    ElseIf vDestination > vDegre Then
        vDegre = vDegre + 0.01
    ElseIf vDestination < vDegre Then
        vDegre = vDegre - 0.01
    End If
    
' Déplace le cercle qui fait office de caméra
    ShpCam.Left = Cos(vDegre) * vRayon + vXRay - (ShpCam.Width / 2)
    ShpCam.Top = -Sin(vDegre) * vRayon + vYRay - (ShpCam.Height / 2)
    
' Déplace le texte indiquand la position en cours
    LblDegre.Left = Cos(vDegre) * (vRayon + 500) + vXRay - (ShpCam.Width / 2)
    LblDegre.Top = -Sin(vDegre) * (vRayon + 500) + vYRay - (ShpCam.Height / 2)
    LblDegre.Caption = Int(vDegre * 180 / 3.141592654 - 180) & "°"
End Sub

Private Sub ClkDelay_Timer() ' Fais office de "Delay"
    If vFonction = 0 Then
' Initialise la position du moteur
        vChaine = "ini;"
        ClkEnvoie.Enabled = True
    End If
    ClkDelay.Enabled = False
End Sub

Private Sub ClkEnvoie_Timer()
Static vCntEnvoie As Integer
Dim vNbCar As Integer
' Ouvre le port série
    MSComm1.PortOpen = True
    
' Définit le nb de caractère de la chaine à envoyer
    vNbCar = Len(vChaine)

' Envoie le caractère en cours
    MSComm1.Output = Mid(vChaine, vCntEnvoie + 1, 1)
    vCntEnvoie = vCntEnvoie + 1
    
' Désactive le timer si on atteint le fin de la chaine
    If vCntEnvoie > vNbCar - 1 Then
        ClkEnvoie.Enabled = False
        vCntEnvoie = 0
    End If
    
' Ferme le port série
    MSComm1.PortOpen = False
End Sub

Private Sub CmdDegre_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vTemp As Integer
Dim vTemp2 As Long
Dim vValeur As Integer
    If Button = 1 Then
        vValeur = Val(CmdDegre(Index).Caption)
        vValTot = vValTot + vValeur
        If vValTot < 0 Or vValTot > 180 Then
            MsgBox "La WebCam ne peut pas aller à cette position!!!", vbCritical, "Erreur"
            vValTot = vValTot - vValeur
        Else
            vDestination = vDegre + vValeur / 180 * 3.141592654
        
            vChaine = "mv" & Left(CmdDegre(Index).Caption, 1) & " " & Abs(Int(Val(CmdDegre(Index).Caption) / vPas)) & ";"
        
            ClkEnvoie.Enabled = True
            ClkCam.Enabled = True
            
            vTemp2 = vDegre * 180 / 3.141592654 + vValeur
            vTemp = vTemp2 / vPas
            vDegText = vTemp * vPas - 180
            
        End If
    ElseIf Button = 2 Then
        vPopup = "C'est le nombre de degré que va tourner la caméra"
        PopupMenu MenuPopup
    End If
End Sub

Private Sub CmdOk_Click()
Dim vTemp As Integer
Dim vTemp2 As Double
    vNbVir = False
    TxtDegre.SetFocus

' Empêche d'entrer des "mauvaises" valeurs
    If TxtDegre.Text = "" Then
        MsgBox "Veuillez entrer une valeur!", vbCritical, "Erreur de saisie"
        TxtDegre.Text = ""
    ElseIf CCur(TxtDegre.Text) < 0 Or CCur(TxtDegre.Text) > 180 Then
        MsgBox "L'angle doit être compris entre 0 et 180°!", vbCritical, "Erreur"
        TxtDegre.Text = ""
    Else
        vValTot = Int(TxtDegre.Text)
        
' Convertit l'angle entré en radian.
        vTemp = CCur(TxtDegre.Text) / vPas
        vDegText = vTemp * vPas
        vDestination = (CInt(TxtDegre.Text) + 180)
        vDestination = vDestination * 3.141592654 / 180
        If vDestination > (360 * 3.141592654 / 180) Then
            vDestination = vDestination - 360 * 3.141592654 / 180
        End If
        ClkCam.Enabled = True
        vTemp2 = Int(((vDegre * 180 / 3.141592654 - 180) - CCur(TxtDegre.Text)) / vPas)
        
' Définit si le moteur va avancer ou reculer et envoie la chaine par le port série
        If vTemp2 < 0 Then
            vChaine = "mv+ " & Abs(vTemp2) & ";"
            ClkEnvoie.Enabled = True
        ElseIf vTemp2 > 0 Then
            vChaine = "mv- " & vTemp2 & ";"
            ClkEnvoie.Enabled = True
        End If
        TxtDegre.Text = ""
    End If
End Sub

Private Sub CmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        vPopup = "Sert à valider le nombre rentré dans le champ de texte"
        PopupMenu MenuPopup
    End If
End Sub

Private Sub Form_Load()
    
' Initialise la vitesse du moteur à 500
    vChaine = "vit 500;"
    ClkEnvoie.Enabled = True
    
' Initialise la position du moteur (voir fonction ClkDelay_Timer)
    vFonction = 0
    ClkDelay.Enabled = True

' Définit la position de départ
    vDegre = 180 * 3.141592654 / 180
    vPas = 3.75
    
' Définit les angles par défaut auxquelles s'arretera la caméra pour prendre une photo
    tArret(0) = 0
    tArret(1) = 90
    tArret(2) = 180
    
' Affiche la Form où l'image de la WebCam apparait
    frmMainCam.Show
    
' Initialise la position de départ des Form pour qu'elles ne se chevauchent pas
    FrmMain.Left = 1000
    FrmMain.Top = 1500
    frmMainCam.Left = 6500
    frmMainCam.Top = 1500
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Décharge la mémoire pour éviter de faire planter le PC
    Unload frmMainCam
    End
End Sub

Private Sub LblDegre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        vPopup = "Indique la position à laquelle se trouve la caméra"
        PopupMenu MenuPopup
    End If
End Sub

Private Sub LblTitre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        vPopup = "Ben... c'est le nom du programme..."
        PopupMenu MenuPopup
    End If
End Sub

Private Sub MenuAidePrinc_Click()
    FrmAide.Show
    FrmMain.Hide
End Sub

Private Sub MenuCapture_Click()
    FrmCapture.Show
    FrmMain.Hide
End Sub

Private Sub MenuCouleur_Click()
' Choisi la couleur de la WebCam
    Dial1.ShowColor
    ShpCam.FillColor = Dial1.Color
    ShpCam.BorderColor = Dial1.Color
End Sub

Private Sub MenuDetail_Click()
    FrmExplorateur.Show
    FrmMain.Hide
End Sub

Private Sub MenuDevisser_Click()
    MsgBox "N'oubliez pas de dessérer le boulon...", vbInformation, "Warning!"
    vChaine = "mv- 500;"
    ClkEnvoie.Enabled = True
End Sub

Private Sub MenuNew_Click()
Dim vCntAct As Integer
    
' Réactive les boutons des degrés
    For vCount = 0 To 5
        CmdDegre(vCount).Enabled = True
    Next
    
' Désactive le Timer qui fait bouger la forme de la caméra...
    ClkCam.Enabled = False
    vDegre = 180 * 3.141592654 / 180
    
'... et ré-initialise sa position...
    ShpCam.Top = 3360
    ShpCam.Left = 1680
    
'... ainsi que celle du label "Degré"
    LblDegre.Left = 1200
    LblDegre.Top = 3360
    LblDegre.Caption = "0°"
    
    TxtDegre.Text = ""
    
' Ré-initialise la position du moteur
    vChaine = "ini;"
    ClkEnvoie.Enabled = True
End Sub

Private Sub MenuPas_Click()
    FrmOption.Show
    FrmMain.Hide
End Sub

Private Sub MenuPopupAide_Click()
    MsgBox vPopup, vbInformation, "Qu'est-ce que c'est ?"
End Sub

Private Sub MenuPropos_Click()
    FrmPropos.Show
    FrmMain.Hide
End Sub

Private Sub MenuQuit_Click()
    Unload frmMainCam
    End
End Sub

Private Sub MenuVisser_Click()
    vChaine = "mv+ 500;"
    ClkEnvoie.Enabled = True
    MsgBox "N'oubliez pas de serrer le boulon...", vbInformation, "Warning"
End Sub

Private Sub MenuVit_Click()
    FrmVit.Show
    FrmMain.Hide
End Sub

Private Sub MenuVoir_Click()
    On Error Resume Next
    frmMainCam.Show
    frmMainCam.Left = 6500
    frmMainCam.Top = 1500
    FrmMain.Left = 1000
    FrmMain.Top = 1500
End Sub

Private Sub TxtDegre_KeyPress(KeyAscii As Integer) ' Protection sur la TextBox
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 And KeyAscii <> 46) Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur de saisie"
        KeyAscii = 0
    End If
    
'Si on presse sur Entre...
    If KeyAscii = 13 Then
        CmdOk_Click
'... ou sur "."
    ElseIf KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
'Si on presse sur ","
    If KeyAscii = 44 Then
        If vNbVir = False Then
            vNbVir = True
            TxtDegre.MaxLength = Len(TxtDegre.Text) + 2
        Else
            KeyAscii = 0
        End If
        If Len(TxtDegre.Text) = 0 Then
            vNbVir = False
            KeyAscii = 0
        End If
    End If
    
'Si on efface une virgule
    If KeyAscii = 8 And Right(TxtDegre.Text, 1) = "," Then
        vNbVir = False
        TxtDegre.MaxLength = 3
    End If
End Sub

Private Sub TxtDegre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        vPopup = "C'est la position à laquelle va se tourner la caméra"
        PopupMenu MenuPopup
    End If
End Sub
