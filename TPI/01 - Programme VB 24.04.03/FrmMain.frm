VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dice Value"
   ClientHeight    =   5070
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VPORTAL2LibCtl.VideoPortal VideoPortal1 
      Height          =   2895
      Left            =   4200
      OleObjectBlob   =   "FrmMain.frx":0000
      TabIndex        =   5
      Top             =   360
      Width           =   4455
   End
   Begin VB.Timer ClkEnCours 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   0
   End
   Begin VB.CommandButton CmdGen 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Timer ClkWebcam 
      Interval        =   1000
      Left            =   2040
      Top             =   3840
   End
   Begin VB.Label LblDiam 
      Caption         =   "LblDiam"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label LblPixTot 
      Caption         =   "Nb pix total : "
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label LblTemps 
      Caption         =   "Temps restant avant la prochaine analyse :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
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
Private Sub ClkEnCours_Timer()
Const FICHIERIMAGE = 1
Const BLANC = 0
Const NOIR = 1
Dim tCount(2), tPixNoir(2), vLong, vInter As Double ' tPixNoir : 0=En cours  1=Max sur une ligne  2=Tot  3=CooX  4=CooY  5=CooX Old  6=CooY Old  7=PosXMax
Dim vCouleur, vChaineTot, vFichier As String
Dim vHaut As Integer
Dim vTemp As Byte
Dim vLast As Boolean    ' False = 0 = Blanc      True = 1 = Noir
Static vCount As Integer

List1.Clear

    vFichier = "c:\image"
    If vCount = 0 Then
        VideoPortal1.PictureToFile 0, 24, vFichier, ""
        vCount = 1
    Else
        vCount = 0
    
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
            Get #FICHIERIMAGE, tCount(0), vTemp
            vCouleur = vTemp & ", " & vCouleur
            tCount(1) = tCount(1) + 1
            If tCount(1) = 4 Then   ' Quand un pixel (3 octets) est lu
                If Val(Left(vCouleur, 3)) < 100 And Val(Mid(vCouleur, 6, 3)) < 50 And Val(Mid(vCouleur, 11, 3)) < 50 Then   'Si il y a un pixel noir
                    tPixNoir(0) = tPixNoir(1) ' Old Y
                    tPixNoir(1) = Int((((tCount(0) - 54) / 3) - 2) / vLong)     ' Coordonnée Y
                    If vLast = BLANC And tPixNoir(1) > (tPixNoir(0) + 1) Then
'                    If vLast = BLANC And (tPixNoir(1) > (tPixNoir(0) + 1) Or tPixNoir(0) = tPixNoir(1)) Then
                        List1.AddItem (tPixNoir(0) & ", " & tPixNoir(1))
                        vInter = vInter + 1
                        vLast = NOIR
                    End If
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
        LblPixTot.Caption = "Nombre de pixel noir : " & tPixNoir(2)
        LblDiam.Caption = "Nombre de points : " & vInter

    '        MsgBox "compté : " & tPixNoir(2), vbInformation

        Kill "c:\image"
        ClkEnCours.Enabled = False
    End If
End Sub

Private Sub ClkWebcam_Timer()
Static vCount As Integer
    LblTemps.Caption = "Temps restant avant la prochaine analyse : " & Int(FrmOption.TxtTemps.Text) - vCount & " sec"
    vCount = vCount + 1
    If vCount = 10 Then
        'appel fonction princ
        vCount = 0
    End If
End Sub

Private Sub CmdGen_Click()
    FrmEnCours.Show
    ClkEnCours.Enabled = True
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    VideoPortal1.PrepareControl "QCSDK_VBDEMO", "HKEY_LOCAL_MACHINE\Software\Logitech\QCSDK_VBDEMO", 0
    
    ' Essaie de connecter une caméra
    If VideoPortal1.ConnectCamera2() = 0 Then
        MsgBox "Impossible de connecter une caméra", vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' Si une caméra est trouvée ça enclanche la prévisualisation de l'image
    VideoPortal1.EnablePreview = 1
End Sub

Private Sub MenuFichierQuitter_Click()
    End
End Sub

Private Sub MenuOutilsOption_Click()
    ClkWebcam.Enabled = False
    FrmOption.Show
End Sub
