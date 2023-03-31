VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dice Value"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7530
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7530
   Begin VB.CommandButton CmdOption 
      Caption         =   "&Option"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   4440
      Width           =   2175
   End
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
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tTabImage(130, 160) As Boolean
Dim tCoo(1, 30) As Integer ' 4 dés maxi * 6 points = 24... -1 = 23

Function AnalyseImage()
Dim vNbTot As Integer
Dim vCount As Integer

    For vCount = 0 To 30        ' Remise à zéro de toutes les cellules
        tCoo(0, vCount) = 0     ' (Nécessaire dès la 2ème fois qu'on appel cette fonction)
        tCoo(1, vCount) = 0
    Next

    LectureImage             ' Le tableau tTabImage contient un 1 aux coordonnées où il y a un pixel noir
    vNbTot = NbPointsTotal   ' vNbTot vaut le nombre de points total
    NbPointsParDes vNbTot    ' Cette fonction sépare la valeur de chaque dés

    If vNbTot <= Int(FrmOption.TxtScore.Text) Then
        MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez gagné!", vbInformation, "Gagné"
    Else
        MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez perdu!", vbInformation, "Perdu"
    End If
End Function

Function LectureImage()
Const FICHIERIMAGE = 1
Const NOIR = 50
Dim vFichier As String
Dim vTemp As Byte
Dim vLargeur As Integer
Dim tCouleur(4) As Byte
Dim vLigne As Integer
Dim vColonne As Integer
Dim tCount(2) As Double

    On Error Resume Next
   ' vFichier = "c:\image.bmp"
   ' VideoPortal1.PictureToFile 0, 24, vFichier, ""
    vFichier = "c:\test.bmp"
    Open vFichier For Random As #FICHIERIMAGE Len = 1   ' Chaque lecture de FICHIERIMAGE  vaudra 1 octet

    '***** Largeur de l'image en pixel *****'
    Get #FICHIERIMAGE, 19, vTemp
    vLargeur = vTemp
    Get #FICHIERIMAGE, 20, vTemp
    vLargeur = vLargeur + 256 * vTemp
    '***************************************'

    '***** Lit le fichier et rempli le tableau en fonction des pixels *****'
    For tCount(0) = 54 To FileLen(vFichier) - 3
        Get #FICHIERIMAGE, tCount(0), vTemp     ' vTemp <-- octet en cours
        tCouleur(3 - tCount(1)) = vTemp
        tCount(1) = tCount(1) + 1      ' Compte le nombre de pixels
        If tCount(1) = 4 Then   ' Quand un pixel (3 octets) est lu
            tCount(2) = tCount(2) + 1
            vLigne = Int(tCount(2)) / (vLargeur) + 1
            vColonne = (tCount(2)) Mod (vLargeur)
            If tCouleur(0) < NOIR And tCouleur(1) < NOIR And tCouleur(2) < NOIR Then 'Si il y a un pixel noir
                tTabImage(vLigne, vColonne) = True
            Else
                tTabImage(vLigne, vColonne) = False
            End If
            tCount(1) = 1
        End If
    Next
    '************************* fin de lecture *****************************'
    Close #FICHIERIMAGE
    Kill vFichier
End Function

Function NbPointsTotal()
Const x = 0
Const y = 1
Dim vCount As Integer
Dim vCount2 As Integer
Dim vNbPtsTot As Integer

    On Error Resume Next
    For vCount = 1 To 120
        For vCount2 = 1 To 160
            If tTabImage(vCount, vCount2) = True And tTabImage(vCount, vCount2 - 1) = False Then
                If IsOver(vCount - 1, vCount2) = False Then
                    tCoo(x, vNbPtsTot) = vCount
                    tCoo(y, vNbPtsTot) = vCount2
                    vNbPtsTot = vNbPtsTot + 1
                End If
            End If
        Next
    Next
    NbPointsTotal = vNbPtsTot
End Function

Function IsOver(ByVal vLigne As Integer, ByVal vColonne As Integer)
Dim vCntOver As Integer
Dim vLimite As Integer
Dim vOver As Boolean

    On Error Resume Next
    If vColonne < 5 Then    ' Empeche les "indice en dehors de la plage"
        vLimite = vColonne
    ElseIf vColonne > 155 Then
        vLimite = 160 - vColonne
    End If

    For vCntOver = vColonne - vLimite To vColonne + vLimite    ' Parcours 5 colonnes avant la ligne du dessus et 5 colonnes après
        If tTabImage(vLigne, vCntOver) = True Then
            vOver = True
            Exit For
        End If
    Next
    IsOver = vOver
End Function

Function NbPointsParDes(ByVal vNbTot As Integer)
Const DISTANCE = 30
Dim vCountDes As Integer
Dim vCountDes2 As Integer
Dim tPointsParDes(20) As Integer ' 3 dés maxi ==> 0 à 2
Dim vNbDif As Integer
Dim vTrop As Integer

    On Error Resume Next
    Do
        vCountDes2 = vCountDes + 1  ' Incrémentation du 1er compteur
       ' vNbTot = vNbTot - vPlus
       ' vPlus = 0
        Do
            If tCoo(0, vCountDes) <> 0 Then     ' Si il y a une coordonnée
                If Abs(tCoo(0, vCountDes) - tCoo(0, vCountDes2)) < DISTANCE And Abs(tCoo(1, vCountDes) - tCoo(1, vCountDes2)) < DISTANCE Then   ' Si 2 coo sont proches
                    If tPointsParDes(vCountDes) = 0 Then
                        tPointsParDes(vCountDes) = 2
                        vNbDif = vNbDif + 1     ' Un dé en plus
                    Else
                        tPointsParDes(vCountDes) = tPointsParDes(vCountDes) + 1     ' Un point en plus sur le dé
                    End If
                    tCoo(0, vCountDes2) = 0 ' Effacement de la coordonnée pour ne pas la compter plusieurs fois
                'ElseIf tPointsParDes(vNbDif) = 0 Then
                '    tPointsParDes(vNbDif) = 1
                '    vNbDif = vNbDif + 1
                End If
                If tPointsParDes(vCountDes) = 0 Then tPointsParDes(vCountDes) = 1 ' S'il n'y a qu'un seul point sur un dé
            End If
            vCountDes2 = vCountDes2 + 1     ' Incrémentation du 2ème compteur
        Loop While (vCountDes2 < vNbTot)

        If tCoo(0, vCountDes) <> 0 Then
            tCoo(0, vCountDes) = 0
        End If
        vCountDes = vCountDes + 1
    Loop While (vCountDes < vNbTot)

    '********** Affichage dans la ListBox **********'
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
            If tPointsParDes(vCountDes) <> 0 Then
                ListDetails.AddItem "Dé n° " & vCountDes + 1 - vTrop & ":     " & tPointsParDes(vCountDes) & " point(s)"
            Else
                vTrop = vTrop + 1
            End If
        Next

        ListDetails.AddItem ""
        ListDetails.AddItem "Total : " & vNbTot & " points"
    End If
    '*************** Fin d'affichage ***************'
End Function

Private Sub ChkTime_Click()
    If ChkTime.Value = Checked Then ClkWebcam.Enabled = True Else ClkWebcam.Enabled = False
End Sub

Private Sub ClkWebcam_Timer()
Static vCountClk As Integer
    If Int(FrmOption.TxtTemps.Text) - vCountClk < 0 Then vCountClk = 0
    LblTemps.Caption = "Temps restant avant la prochaine analyse : " & Int(FrmOption.TxtTemps.Text) - vCountClk & " sec"
    vCountClk = vCountClk + 1
    If vCountClk = Int(FrmOption.TxtTemps.Text) + 1 Then
        If ChkTime.Value = Checked Then
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

    ' Cherche les drivers de la webcam
    If VideoPortal1.ConnectCamera2() = 0 Then
        MsgBox "Les drivers de la webcam ne sont pas installés.", vbCritical, "Erreur"
 '       End
    End If

    ' Essaie de connecter une caméra
    If VideoPortal1.ConnectCamera(0) = 0 Then
        MsgBox "Impossible de connecter une webcam, vérifiez qu'elle est bien branchée", vbCritical, "Erreur"
  '      End
    End If

    ' Enclanche la prévisualisation de l'image
    VideoPortal1.EnablePreview = 1
  '  VideoPortal1.SetCameraPropertyLong PROPERTY_MOTION_DETECTION_MODE, 1  'Necessaire pour la détéction de mouvements
End Sub

Private Sub CmdOption_Click()
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

Private Sub ClkLance_Timer()
    If PBMove.Value = 0 Then
        If ChkMov.Value = Checked Then
            AnalyseImage
        End If
        ClkLance.Enabled = False
    End If
End Sub
