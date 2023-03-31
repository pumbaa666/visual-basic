VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dice Value"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   7530
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7530
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
      Caption         =   "D�tecteur de mouvements"
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
Option Explicit
Dim tTabImage(200, 200) As Integer
Dim tCoo(1, 17) As Integer ' 3 d�s * 6 points = 18... -1 = 17

Function AnalyseImage()
Dim vNbTot As Integer

    FrmEnCours.Show
    
    LectureImage             ' Le tableau tTabImage contient un 1 aux coordonn�es o� il y a un pixel noir
    vNbTot = NbPointsTotal   ' vNbTot vaut le nombre de points total
    NbPointsParDes vNbTot

    FrmEnCours.Hide
   ' NbPointsParDes vNbPointsTot, tAll
   ' If vNbPointsTot <= Int(FrmOption.TxtScore.Text) Then
   '     MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez gagn�!", vbInformation, "Gagn�"
   ' Else
   '     MsgBox "Il fallait faire " & Int(FrmOption.TxtScore.Text) & " point(s) maximum pour gagner" & Chr$(13) & "Vous avez perdu!", vbInformation, "Perdu"
   ' End If
   ' Kill "c:\image"
End Function

Function NbPointsParDes(ByVal vNbTot As Double)
Dim vCountDes As Integer
Dim vCountDes2 As Integer
Dim tPointsParDes(20) As Integer ' 3 d�s maxi ==> 0 � 2
Dim vPlus As Integer
Dim vNbDif As Integer
Dim vTrop As Integer
    
    Do
        vCountDes2 = vCountDes + 1  ' Incr�mentation du 1er compteur
       ' vNbTot = vNbTot - vPlus
       ' vPlus = 0
        Do
            If tCoo(0, vCountDes) <> 0 Then     ' Si il y a une coordonn�e
                If Abs(tCoo(0, vCountDes) - tCoo(0, vCountDes2)) < 40 And Abs(tCoo(1, vCountDes) - tCoo(1, vCountDes2)) < 40 Then   ' Si 2 coo sont proches
                    If tPointsParDes(vCountDes) = 0 Then
                        tPointsParDes(vCountDes) = 2
                        vNbDif = vNbDif + 1     ' Un d� en plus
                    Else
                        tPointsParDes(vCountDes) = tPointsParDes(vCountDes) + 1     ' Un point en plus sur le d�
                    End If
                    tCoo(0, vCountDes2) = 0 ' Effacement de la coordonn�e pour ne pas la compter plusieurs fois
       '             vPlus = vPlus + 1
                'ElseIf tPointsParDes(vNbDif) = 0 Then
                '    tPointsParDes(vNbDif) = 1
                '    vNbDif = vNbDif + 1
                End If
                If tPointsParDes(vCountDes) = 0 Then tPointsParDes(vCountDes) = 1
            End If
            vCountDes2 = vCountDes2 + 1     ' Incr�mentation du 2�me compteur
        Loop While (vCountDes2 < vNbTot)

        If tCoo(0, vCountDes) <> 0 Then
            vPlus = vPlus + 1
            tCoo(0, vCountDes) = 0
        End If
        vCountDes = vCountDes + 1
    Loop While (vCountDes < vNbTot)

    ListDetails.Clear
    If vNbDif = 0 Then
        ListDetails.AddItem "Il n'y a pas de d�s"
    ElseIf vNbDif = 1 Then
        ListDetails.AddItem "Il y a 1 d�"
        ListDetails.AddItem ""
        ListDetails.AddItem "Valeur : " & tPointsParDes(0) & " point(s)"
    Else
        ListDetails.AddItem "Il y a " & vNbDif & " d�s"
        ListDetails.AddItem ""

        For vCountDes = 0 To 20
            If tPointsParDes(vCountDes) <> 0 Then ListDetails.AddItem "D� n� " & vCountDes + 1 - vTrop & ":     " & tPointsParDes(vCountDes) & " point(s)" Else vTrop = vTrop + 1
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
    FrmImage.Show
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

Function LectureImage()
Const FICHIERIMAGE = 1
Dim vFichier As String
Dim vTemp As Byte
Dim vLargeur As Integer
Dim tCouleur(4) As Integer
Dim vLigne As Integer
Dim vColonne As Integer
Dim vNbPixel As Integer
Dim tCount(2) As Double
Dim vLongFile As Double

 '   vFichier = "c:\image"
 '   VideoPortal1.PictureToFile 0, 24, vFichier, ""
    vFichier = "c:\test.bmp"
    Open vFichier For Random As #FICHIERIMAGE Len = 1
    vLongFile = FileLen(vFichier)
    
    Get #FICHIERIMAGE, 19, vTemp
    vLargeur = vTemp
    Get #FICHIERIMAGE, 20, vTemp
    vLargeur = vLargeur + 256 * vTemp

    ' Commence la lecture du fichier et rempli le tableau en fonction des pixels
    For tCount(0) = 54 To vLongFile - 3
        Get #FICHIERIMAGE, tCount(0), vTemp     ' vTemp <-- octet en cours
        tCouleur(3 - tCount(1)) = vTemp
        tCount(1) = tCount(1) + 1
        If tCount(1) = 4 Then   ' Quand un pixel (3 octets) est lu
            vNbPixel = vNbPixel + 1
            ' R�cup�ration des pixels pour les mettre dans un tableau � 2 dimensions
            vLigne = 1 + Int(vNbPixel) / (vLargeur)
            vColonne = (vNbPixel) Mod (vLargeur)

            If tCouleur(0) < 70 And tCouleur(1) < 70 And tCouleur(2) < 70 Then 'Si il y a un pixel noir
                tTabImage(vLigne, vColonne) = 1
                FrmImage.PSet (vColonne * 15 + 200, vLigne * 15 + 200)
            End If
            tCount(1) = 1
        End If
    Next
    Close #FICHIERIMAGE
End Function

Function NbPointsTotal()
Const X = 0
Const Y = 1
Dim vCount As Integer
Dim vCount2 As Integer
Dim vNbPtsTot As Integer
Dim tPixNoir(1) As Double

    For vCount = 1 To 120
        For vCount2 = 1 To 160
            If tTabImage(vCount, vCount2) = 1 And tTabImage(vCount, vCount2 - 1) = 0 Then
                If IsOver(vCount - 1, vCount2) = False Then
                    tCoo(X, vNbPtsTot) = vCount
                    tCoo(Y, vNbPtsTot) = vCount2
                    vNbPtsTot = vNbPtsTot + 1
                End If
            End If
        Next
    Next
    NbPointsTotal = vNbPtsTot
End Function

Function IsOver(ByVal vLigne As Integer, ByVal vColonne As Integer)
Dim vCntOver As Integer
Dim vOver As Boolean

    For vCntOver = vColonne - 5 To vColonne + 5
        If tTabImage(vLigne, vCntOver) = 1 Then
            vOver = True
            Exit For
        End If
    Next
    IsOver = vOver
End Function

Function EcrireImage(ByVal fCount As Double)
Dim vCouleur As Byte
    Open "c:\blanc.bmp" For Random As #1
    vCouleur = 0
    Put #1, fCount - 2, vCouleur
    vCouleur = 255
    Put #1, fCount - 1, vCouleur
    vCouleur = 0
    Put #1, fCount, vCouleur
    Close #1
End Function
