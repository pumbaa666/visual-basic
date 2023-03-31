VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dice Value"
   ClientHeight    =   5760
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7530
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkLance 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4440
      Top             =   4680
   End
   Begin VB.ListBox ListDetails 
      Height          =   3570
      Left            =   5160
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar PBMove 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CheckBox ChkMov 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   4920
      Width           =   255
   End
   Begin VB.CheckBox ChkTime 
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VPORTAL2LibCtl.VideoPortal VideoPortal1 
      Height          =   3555
      Left            =   240
      OleObjectBlob   =   "FrmMain.frx":0442
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Timer ClkWebcam 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "Détecteur de mouvements"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   2055
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
Const X = 0
Const Y = 1

Dim vLast As Boolean    ' False = 0 = Blanc      True = 1 = Noir

Dim tCouleur(3), vTemp As Byte

Dim vLastMeme As Integer
Dim vLigneEnCours As Integer
Dim tAll(1, 170) As Integer ' 3 dés * 6 points = 18... -1 = 17

Dim tCount(2) As Double
Dim tPixNoir(1) As Double
Dim vLong As Double
Dim vNbPointsTot As Double ' 0 = Old Y  1 = Y en cours

Dim vFichier As String

    vFichier = "c:\image"
    VideoPortal1.PictureToFile 0, 24, vFichier, ""
    ' vLong contiendra la largeur de l'image en pixel
    vFichier = "c:\blanc2.bmp"
    Open vFichier For Random As #FICHIERIMAGE Len = 1
    Get #FICHIERIMAGE, 19, vTemp
    vLong = vTemp
    Get #FICHIERIMAGE, 20, vTemp
    vLong = vLong + 256 * vTemp

    ' Commence la lecture du fichier
    For tCount(0) = 54 To FileLen(vFichier)
        If ((tCount(0) - 54) - 2) Mod (vLong * 3) = 0 Then vLigneEnCours = vLigneEnCours + 1 ' Si on passe à la ligne suivante
        Get #FICHIERIMAGE, tCount(0), vTemp     ' vTemp <-- octet en cours
        tCouleur(3 - tCount(1)) = vTemp
        tCount(1) = tCount(1) + 1
        If tCount(1) = 4 Then   ' Quand un pixel (3 octets) est lu
            If tCouleur(0) < 70 And tCouleur(1) < 70 And tCouleur(2) < 70 Then 'Si il y a un pixel noir
            '    ListDetails.AddItem tCouleur(0) & ", " & tCouleur(1) & ", " & tCouleur(2)
                tPixNoir(0) = tPixNoir(1) ' Old Y
                tPixNoir(1) = Int((((tCount(0) - 54) / 3) - 2) / vLong)     ' Coordonnée Y
                If vLastMeme <> 0 Then vLastMeme = vLigneEnCours
                If vLast = BLANC Then
                    If tPixNoir(1) > (tPixNoir(0) + 1) Then     ' Si le pixel noir suivant est 2 lignes plus loin que celui d'avant. Si c'est le cas, ça veut dire que c'est un nouveau point
                        tAll(X, vNbPointsTot) = Int((((tCount(0) - 54) / 3) - 2) / vLong)   ' Sauvegarde de façon définitive la coordonnée X du point
                        tAll(Y, vNbPointsTot) = ((tCount(0) - 54) / 3) - (tAll(X, vNbPointsTot) * vLong) ' Sauvegarde de façon définitive la coordonnée Y du point
                        vNbPointsTot = vNbPointsTot + 1
                        vLastMeme = 0
                    ElseIf tPixNoir(1) = tPixNoir(0) And (tPixNoir(1) > vLastMeme + 1) Then ' Or tPixNoir(1) = vLastMeme) Then   ' Sinon, si il est sur la même ligne et qu'il n'a pas encore été compté...
                        tAll(X, vNbPointsTot) = Int((((tCount(0) - 54) / 3) - 2) / vLong)   ' Sauvegarde de façon définitive la coordonnée X du point
                        tAll(Y, vNbPointsTot) = (tCount(0) - 54) / 3 - tAll(X, vNbPointsTot) * vLong     ' Sauvegarde de façon définitive la coordonnée Y du point
                        vNbPointsTot = vNbPointsTot + 1
                        'vLastMeme = 0
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
    PointsParDes vNbPointsTot, tAll
   ' If vNbPointsTot <= Int(FrmOption.TxtScore.Text) Then
   '     MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez gagné!", vbInformation, "Gagné"
   ' Else
   '     MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez perdu!", vbInformation, "Perdu"
   ' End If
   ' Kill "c:\image"
    
'    ListDetails.Clear
'    For vCount = 0 To 23
'        ListDetails.AddItem tAll(0, vCount)
'    Next
End Function

Function PointsParDes(ByVal vNbTot As Double, ByVal tAllPar)
Dim vCountDes As Integer
Dim vCountDes2 As Integer
Dim tPointsParDes(20) As Integer ' 3 dés maxi ==> 0 à 2
Dim vPlus As Integer
Dim vNbDif As Integer
Dim vTrop As Integer
    Do
        vCountDes2 = vCountDes + 1  ' Incrémentation du 1er compteur
       ' vNbTot = vNbTot - vPlus
       ' vPlus = 0
        Do
            If tAllPar(0, vCountDes) <> 0 Then     ' Si il y a une coordonnée
                If Abs(tAllPar(0, vCountDes) - tAllPar(0, vCountDes2)) < 40 And Abs(tAllPar(1, vCountDes) - tAllPar(1, vCountDes2)) < 40 Then   ' Si 2 coo sont proches
                    If tPointsParDes(vCountDes) = 0 Then
                        tPointsParDes(vCountDes) = 2
                        vNbDif = vNbDif + 1     ' Un dé en plus
                    Else
                        tPointsParDes(vCountDes) = tPointsParDes(vCountDes) + 1     ' Un point en plus sur le dé
                    End If
                    tAllPar(0, vCountDes2) = 0 ' Effacement de la coordonnée pour ne pas la compter plusieurs fois
       '             vPlus = vPlus + 1
                'ElseIf tPointsParDes(vNbDif) = 0 Then
                '    tPointsParDes(vNbDif) = 1
                '    vNbDif = vNbDif + 1
                End If
                If tPointsParDes(vCountDes) = 0 Then tPointsParDes(vCountDes) = 1
            End If
            vCountDes2 = vCountDes2 + 1     ' Incrémentation du 2ème compteur
        Loop While (vCountDes2 < vNbTot)

        If tAllPar(0, vCountDes) <> 0 Then
            vPlus = vPlus + 1
            tAllPar(0, vCountDes) = 0
        End If
        vCountDes = vCountDes + 1
    Loop While (vCountDes < vNbTot)

    ListDetails.Clear
    If vNbDif = 0 Then
        ListDetails.AddItem "Il n'y a pas de dés"
    ElseIf vNbDif = 1 Then
        ListDetails.AddItem "Il y a 1 dé"
        ListDetails.AddItem ""
        ListDetails.AddItem "Valeur : " & tPointsParDes(0) & " point(s)"
    Else
        ListDetails.AddItem "Il y a " & vNbDif & " dés"
        ListDetails.AddItem ""

        For vCountDes = 0 To 20
            If tPointsParDes(vCountDes) <> 0 Then ListDetails.AddItem "Dé n° " & vCountDes + 1 - vTrop & ":     " & tPointsParDes(vCountDes) & " point(s)" Else vTrop = vTrop + 1
        Next

        ListDetails.AddItem ""
        ListDetails.AddItem "Total : " & vNbTot & " points"
    End If
End Function

Private Sub ChkTime_Click()
    If ChkTime.Value = Checked Then ClkWebcam.Enabled = True Else ClkWebcam.Enabled = False
End Sub

Private Sub ClkLance_Timer()
    If PBMove.Value = 0 Then
        If ChkMov.Value = Checked Then
            FrmEnCours.Show
            AnalyseImage
        End If
        ClkLance.Enabled = False
    End If
End Sub

Private Sub ClkWebcam_Timer()
Static vCountClk As Integer
    If Int(FrmOption.TxtTemps.Text) - vCountClk < 0 Then vCountClk = 0
    LblTemps.Caption = "Temps restant avant la prochaine analyse : " & Int(FrmOption.TxtTemps.Text) - vCountClk & " sec"
    vCountClk = vCountClk + 1
    If vCountClk = Int(FrmOption.TxtTemps.Text) + 1 Then
        If ChkTime.Value = Checked Then
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
    VideoPortal1.EnableUIElements UIELEMENT_STATUSBAR, 0, 1 ' Necessaire pour la détéction de mouvements

    ' Essaie de connecter une caméra
    If VideoPortal1.ConnectCamera2() = 0 Then
        MsgBox "Impossible de connecter une caméra", vbCritical, "Erreur"
        Exit Sub
    End If

    ' Si une caméra est trouvée ça enclanche la prévisualisation de l'image
    VideoPortal1.EnablePreview = 1
    VideoPortal1.SetCameraPropertyLong PROPERTY_MOTION_DETECTION_MODE, 1  'Necessaire pour la détéction de mouvements

    AnalyseImage
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
        If lParam1 > 5 Then ClkLance.Enabled = True
    End If
End Sub
