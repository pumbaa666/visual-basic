VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dice Value"
   ClientHeight    =   6645
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5295
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PBMove 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   5160
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
      Height          =   3555
      Left            =   240
      OleObjectBlob   =   "FrmMain.frx":0442
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Timer ClkWebcam 
      Interval        =   1000
      Left            =   3480
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "D�tecteur de mouvements"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
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
Dim tAll(1, 23) As Integer ' 4 d�s * 6 points : 24... -1 = 23

Function AnalyseImage()
Const FICHIERIMAGE = 1
Const BLANC = 0
Const NOIR = 1
Const X = 0
Const Y = 1

Dim vLast As Boolean    ' False = 0 = Blanc      True = 1 = Noir

Dim tCouleur(3), vTemp As Byte

Dim vLastMeme As Integer
Dim vLigneEnCours As Integer

Dim tCount(2) As Double
Dim tPixNoir(1) As Double
Dim vLong As Double
Dim vNbPointsTot As Double ' 0=Old Y  1=En cours

Dim vFichier As String

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
        If ((tCount(0) - 54) - 2) Mod (vLong * 3) = 0 Then  ' Si on passe � la ligne suivante
            vLigneEnCours = vLigneEnCours + 1
        End If
        Get #FICHIERIMAGE, tCount(0), vTemp     ' vTemp <-- octet en cours
        tCouleur(3 - tCount(1)) = vTemp
        tCount(1) = tCount(1) + 1
        If tCount(1) = 4 Then   ' Quand un pixel (4 octets) est lu
            If tCouleur(0) < 40 And tCouleur(1) < 40 And tCouleur(2) < 40 Then 'Si il y a un pixel noir
                'List1.AddItem tCouleur(0) & ", " & tCouleur(1) & ", " & tCouleur(2)
                tPixNoir(0) = tPixNoir(1) ' Old Y
                tPixNoir(1) = Int((((tCount(0) - 54) / 3) - 2) / vLong)     ' Coordonn�e Y
                If vLastMeme <> 0 Then
                    vLastMeme = vLigneEnCours
                End If
                If vLast = BLANC Then
                    If tPixNoir(1) > (tPixNoir(0) + 1) Then     ' Si le pixel noir suivant est 2 lignes plus loin que celui d'avant. Si c'est le cas, �a veut dire que c'est un nouveau point
'                        EcrireImage tCount(0), 111
                        tAll(X, vNbPointsTot) = Int((((tCount(0) - 54) / 3) - 2) / vLong)   ' Sauvegarde de fa�on d�finitive la coordonn�e X du point
                        tAll(Y, vNbPointsTot) = ((tCount(0) - 54) / 3) - (tAll(X, vNbPointsTot) * vLong) ' Sauvegarde de fa�on d�finitive la coordonn�e Y du point
                        vNbPointsTot = vNbPointsTot + 1
                        vLastMeme = 0
                    ElseIf tPixNoir(1) = tPixNoir(0) And tPixNoir(1) > vLastMeme + 1 Then   ' Sinon, si il est sur la m�me ligne et qu'il n'a pas encore �t� compt�...
'                        EcrireImage tCount(0), 179
                        tAll(X, vNbPointsTot) = Int((((tCount(0) - 54) / 3) - 2) / vLong)   ' Sauvegarde de fa�on d�finitive la coordonn�e X du point
                        tAll(Y, vNbPointsTot) = (tCount(0) - 54) / 3 - tAll(X, vNbPointsTot) * vLong     ' Sauvegarde de fa�on d�finitive la coordonn�e Y du point
                        vNbPointsTot = vNbPointsTot + 1
                        vLastMeme = vLigneEnCours
                    End If
                End If
                vLast = NOIR
            Else
                vLast = BLANC
            End If
            tCount(1) = 1
        End If
    Next
    Close #FICHIERIMAGE
    ' Fin de la lecture

    FrmEnCours.Hide
    LblDiam.Caption = "Nombre de points : " & vNbPointsTot
    PointsParDes (vNbPointsTot)
    If vNbPointsTot <= Int(FrmOption.TxtScore.Text) Then
        MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez gagn�!", vbInformation, "Gagn�"
    Else
        MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez perdu!", vbInformation, "Perdu"
    End If
   ' Kill "c:\image"
    
'    List1.Clear
'    For vCount = 0 To 23
'        List1.AddItem tAll(0, vCount)
'    Next
End Function

Function PointsParDes(ByVal vNbTot)
Dim vCountDes As Integer
Dim vCountDes2 As Integer
Dim tPointsParDes(8) As Integer ' 9 d�s maxi ==> 0 � 8

'    For vCountDes = 0 To vNbTot - 1
'        tPointsParDes(vCountDes) = 1
'    Next

    For vCountDes = 0 To vNbTot - 1
        For vCountDes2 = vCountDes + 1 To vNbTot - 1
            If Abs(tAll(1, vCountDes) - tAll(1, vCountDes2)) < 40 And tAll(1, vCountDes) <> 0 Then
                If tPointsParDes(vCountDes) = 0 Then
                    tPointsParDes(vCountDes) = 2
                Else
                    tPointsParDes(vCountDes) = tPointsParDes(vCountDes) + 1
                End If
                tAll(1, vCountDes2) = 0
'                vNbTot = vNbTot - 1
            ElseIf tPointsParDes(vCountDes) = 0 Then 'And tAll(1, vCountDes) <> 0 Then
                tPointsParDes(vCountDes) = 1
            End If
        Next
        tAll(1, vCountDes) = 0
    Next
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
Static vCountClk As Integer
    If Int(FrmOption.TxtTemps.Text) - vCountClk < 0 Then
        vCountClk = 0
    End If
    LblTemps.Caption = "Temps restant avant la prochaine analyse : " & Int(FrmOption.TxtTemps.Text) - vCountClk & " sec"
    vCountClk = vCountClk + 1
    If vCountClk = Int(FrmOption.TxtTemps.Text) + 1 Then
        If Check1.Value = Checked Then
            FrmEnCours.Show
            AnalyseImage
        End If
        vCountClk = 0
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    VideoPortal1.PrepareControl "QCSDK_VBDEMO", "HKEY_LOCAL_MACHINE\Software\Logitech\QCSDK_VBDEMO", 0
    VideoPortal1.EnableUIElements UIELEMENT_STATUSBAR, 0, 1 ' Necessaire pour la d�t�ction de mouvements

    ' Essaie de connecter une cam�ra
    If VideoPortal1.ConnectCamera2() = 0 Then
        MsgBox "Impossible de connecter une cam�ra", vbCritical, "Erreur"
        Exit Sub
    End If

    ' Si une cam�ra est trouv�e �a enclanche la pr�visualisation de l'image
    VideoPortal1.EnablePreview = 1
    VideoPortal1.SetCameraPropertyLong PROPERTY_MOTION_DETECTION_MODE, 1  'Necessaire pour la d�t�ction de mouvements

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
