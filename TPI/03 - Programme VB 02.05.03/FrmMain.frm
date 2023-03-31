VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dice Value"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5340
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PBMove 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Timer ClkLance 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   4560
   End
   Begin VPORTAL2LibCtl.VideoPortal VideoPortal1 
      Height          =   2895
      Left            =   240
      OleObjectBlob   =   "FrmMain.frx":0442
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5640
      Width           =   4815
   End
   Begin VB.Timer ClkWebcam 
      Interval        =   1000
      Left            =   3240
      Top             =   480
   End
   Begin VB.Label LblDiam 
      Caption         =   "LblDiam"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label LblTemps 
      Caption         =   "Temps restant avant la prochaine analyse :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu MenuOutils 
      Caption         =   "Outils"
      Begin VB.Menu MenuOutilsOption 
         Caption         =   "Option"
      End
   End
   Begin VB.Menu MenuAide 
      Caption         =   "Aide"
      Begin VB.Menu MenuAideAbout 
         Caption         =   "A propos"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function AnalyseImage()
Const FICHIERIMAGE = 1
Const BLANC = 0
Const NOIR = 1
Dim tCount(2), tPixNoir(1), vLong, vNbPoints As Double ' tPixNoir : 0=Old Y  1=En cours
Dim vCouleur, vChaineTot, vFichier As String
Dim vHaut, vLastMeme As Integer
Dim vTemp As Byte
Dim vLast As Boolean    ' False = 0 = Blanc      True = 1 = Noir

    vFichier = "c:\image"
    VideoPortal1.PictureToFile 0, 24, vFichier, ""
    ' vLong contiendra la largeur de l'image en pixel
    vFichier = "c:\blanc.bmp"
    Open vFichier For Random As #FICHIERIMAGE Len = 1
    Get #FICHIERIMAGE, 19, vTemp
    vLong = vTemp
    Get #FICHIERIMAGE, 20, vTemp
    vLong = vLong + 256 * vTemp

    ' Commencement de la lecture du fichier
    tCount(0) = 54
    For tCount(0) = 54 To FileLen(vFichier)
        Get #FICHIERIMAGE, tCount(0), vTemp     ' vTemp <-- octet en cours
        vCouleur = vTemp & ", " & vCouleur
        tCount(1) = tCount(1) + 1
        If tCount(1) = 4 Then   ' Quand un pixel (3 octets) est lu
            If Val(Left(vCouleur, 3)) < 100 And Val(Mid(vCouleur, 6, 3)) < 10 And Val(Mid(vCouleur, 11, 3)) < 10 Then   'Si il y a un pixel noir
                tPixNoir(0) = tPixNoir(1) ' Old Y
                tPixNoir(1) = Int((((tCount(0) - 54) / 3) - 2) / vLong)     ' Coordonnée Y
                If vLastMeme <> 0 Then
                    vLastMeme = vLastMeme + 1
                End If
                If vLast = BLANC Then
                    If tPixNoir(1) > (tPixNoir(0) + 1) Then
                        EcrireImage tCount(0), 111
                        vNbPoints = vNbPoints + 1
                    '    vLastMeme = tPixNoir(1)
                    ElseIf tPixNoir(1) = tPixNoir(0) And tPixNoir(1) > vLastMeme + 1 Then
                        EcrireImage tCount(0), 179
                        vNbPoints = vNbPoints + 1
                        vLastMeme = tPixNoir(1)
                    End If
                End If
                vLast = NOIR
            Else
                vLast = BLANC
            End If
            vCouleur = ""
            tCount(1) = 1
        End If
    Next
    Close #FICHIERIMAGE
    ' Fin de la lecture

    FrmEnCours.Hide
    LblDiam.Caption = "Nombre de points : " & vNbPoints
    If vNbPoints <= Int(FrmOption.TxtScore.Text) Then
        MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez gagné!", vbInformation, "Gagné"
    Else
        MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez perdu", vbInformation, "Perdu"
    End If
    Kill "c:\image"
End Function

Private Sub ClkLance_Timer()
    If PBMove.Value = 0 Then
        If Check2.Value = Checked Then
            FrmEnCours.Show
            AnalyseImage
        End If
        ClkLance.Enabled = False
    End If
End Sub

Private Sub ClkWebcam_Timer()
Static vCount As Integer
    If Int(FrmOption.TxtTemps.Text) - vCount < 0 Then
        vCount = 0
    End If
    LblTemps.Caption = "Temps restant avant la prochaine analyse : " & Int(FrmOption.TxtTemps.Text) - vCount & " sec"
    vCount = vCount + 1
    If vCount = Int(FrmOption.TxtTemps.Text) + 1 Then
        If Check1.Value = Checked Then
            FrmEnCours.Show
            AnalyseImage
        End If
        vCount = 0
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    VideoPortal1.PrepareControl "QCSDK_VBDEMO", "HKEY_LOCAL_MACHINE\Software\Logitech\QCSDK_VBDEMO", 0
    VideoPortal1.EnableUIElements UIELEMENT_STATUSBAR, 0, 1 ' Necessaire pour la détéction de mouvements
    
    ' Essaie de connecter une caméra
    If VideoPortal1.ConnectCamera2() = 0 Then
        MsgBox "Impossible de connecter une caméra", vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' Si une caméra est trouvée ça enclanche la prévisualisation de l'image
    VideoPortal1.EnablePreview = 1
    VideoPortal1.SetCameraPropertyLong PROPERTY_MOTION_DETECTION_MODE, 1  'Necessaire pour la détéction de mouvements
    
   ' AnalyseImage
End Sub

Private Sub MenuAideAbout_Click()
    FrmAbout.Show
    FrmMain.Hide
End Sub

Private Sub MenuFichierQuitter_Click()
    End
End Sub

Private Sub MenuOutilsOption_Click()
    ClkWebcam.Enabled = False
    FrmOption.Show
    FrmMain.Hide
End Sub

Private Sub VideoPortal1_PortalNotification(ByVal lMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParam3 As Long)
    If lMsg = NOTIFICATIONMSG_MOTION Then
        PBMove.Value = lParam1
        If lParam1 > 5 Then
            ClkLance.Enabled = True
        End If
    End If
End Sub

Function EcrireImage(ByVal fCount As Double, ByVal vCouleur As Byte)
    Put #1, fCount - 2, vCouleur
    Put #1, fCount - 1, vCouleur
    Put #1, fCount, vCouleur
End Function
