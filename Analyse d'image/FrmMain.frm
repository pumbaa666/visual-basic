VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorateur"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "&Raffraichir"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Timer ClkEnCours 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1320
      Top             =   3000
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3000
      Pattern         =   "*.bmp"
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label LblTest 
      Caption         =   "Choisissez l'image à analyser"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuColler 
         Caption         =   "Coller"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuCopier 
         Caption         =   "Copier"
      End
      Begin VB.Menu MenuCouper 
         Caption         =   "Couper"
      End
      Begin VB.Menu Tiret1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDel 
         Caption         =   "Supprimer"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNomFichier As String

Private Sub ClkEnCours_Timer()
    Static vCount As Integer
    If vCount = 0 Then
        FrmEnCours.Show
        FrmEnCours.Top = FrmMain.Top
        FrmEnCours.Left = FrmMain.Left
        FrmMain.Hide
        vCount = vCount + 1
    Else
        Analyse
        ClkEnCours.Enabled = False
        vCount = 0
    End If
End Sub

Private Sub CmdOk_Click()
    ClkEnCours.Enabled = True
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Function Analyse()
Const FICHIERIMAGE = 1
Dim tCount(2), tPixNoir(7), vLong As Double ' tPixNoir : 0=En cours  1=Max  2=Tot  3=CooX  4=CooY  5=CooX Old  6=CooY Old  7=PosXMax
Dim vCouleur, vChaineTot, vFichier, vCouleurAffiche, vPos As String
Dim vSaut, vHaut, tCote(1), tCercle(1), tMax(4), tCoin(3, 1) As Integer  ' tCoin : 0=1er pixel 3=Dernier pixel  tMax : 0=X  1=Y  2=Test   3=Old Plus ou Moins    4=Plus ou Moins
Dim vTemp As Byte

    If File1.FileName = "" Then
        MsgBox "Veuillez choisire une image", vbCritical, "Erreur"
    Else

        ' Met le nom complet du fichier dans vFichier. Il faut faire se test, car quand on est à la racine du disque, il y a automatiquement un \ dans Dir1.Path
        If Len(Dir1.Path) = 3 Then
            vFichier = Dir1.Path & File1.FileName
        Else
            vFichier = Dir1.Path & "\" & File1.FileName
        End If

        ' vLong contiendra la largeur de l'image en pixel
        Open vFichier For Random As #FICHIERIMAGE Len = 1
        Get #FICHIERIMAGE, 19, vTemp
        vLong = vTemp
        Get #FICHIERIMAGE, 20, vTemp
        vLong = vLong + 256 * vTemp

        ' Définit la valeur du saut à faire. (A cause du format BMP, il faut que la largeur de l'image soit un multiple de 4)
        If vLong Mod 4 = 0 Then
            vSaut = 0
        ElseIf vLong Mod 3 = 0 Then
            vSaut = 3
        ElseIf vLong Mod 2 = 0 Then
            vSaut = 2
        Else
            vSaut = 1
        End If

        ' Commencement de la lecture du fichier
        For tCount(0) = 54 To FileLen(vFichier)
            If tCount(0) = 54 + vLong * (3 + tCount(2)) Then ' Si on arrive à la fin d'une ligne
                tCount(0) = tCount(0) + vSaut
                tCount(2) = tCount(2) + vSaut + (4 - vSaut)
            End If

            Get #FICHIERIMAGE, tCount(0), vTemp
            vCouleur = vTemp & ", " & vCouleur
            tCount(1) = tCount(1) + 1
            If tCount(1) = 4 Then   ' Quand un pixel (4 octets) est lu
                If Val(Left(vCouleur, 3)) < 100 And Val(Mid(vCouleur, 6, 3)) < 100 And Val(Mid(vCouleur, 11, 3)) < 100 Then   'Si il y a un pixel noir

                    tPixNoir(0) = tPixNoir(0) + 1
                    tPixNoir(2) = tPixNoir(2) + 1

                    tPixNoir(5) = tPixNoir(3)
                    tPixNoir(6) = tPixNoir(4)

                    tPixNoir(4) = Int((((tCount(0) - 54) / 3) - 2) / vLong)     ' Coordonnée X
                    tPixNoir(3) = (tCount(0) - 54) / 3 - tPixNoir(4) * vLong    ' Coordonnée Y

                    If tCercle(0) = 0 Then
                        tCercle(0) = tPixNoir(4)    ' Coordonné Y du 1er pixel du fichier (pour le cercle)
                    End If
                    tCercle(1) = tPixNoir(4)        ' Coordonné Y du dernier pixel du fichier (pour le cercle)

                    If tPixNoir(3) <> tPixNoir(5) And tPixNoir(4) <> tPixNoir(6) Then   ' Pour ne pas lire 2X le même pixel
                        If tCoin(0, 0) = 0 And tCoin(0, 1) = 0 Then
                            tCoin(0, 0) = tPixNoir(3)   ' Coordonné X du 1er pixel du fichier (pour le rectangle)
                            tCoin(0, 1) = tPixNoir(4)   ' Coordonné Y du 1er pixel du fichier (pour le rectangle)
                        End If
                        tCoin(2, 0) = tPixNoir(3)   ' Coordonné X du dernier pixel du fichier (pour le rectangle)
                        tCoin(2, 1) = tPixNoir(4)   ' Coordonné Y du dernier pixel du fichier (pour le rectangle)
                        tMax(3) = tMax(0)

                        ' Pour savoir quand on atteint un coin
                        If tMax(2) = 0 Then
                            tMax(0) = tPixNoir(3)
                            tMax(1) = tPixNoir(4)
                            tMax(2) = 1
                        ElseIf tMax(2) = 1 Then
                            tMax(4) = Sgn(tPixNoir(3) - tMax(0))
                            tMax(2) = 2
                        ElseIf tMax(2) <> 3 Then
                            tMax(3) = tMax(0)
                            tMax(0) = tPixNoir(3)
                            tMax(1) = tPixNoir(4)
                            If (tMax(4) = 1 And tMax(0) < tMax(3)) Or (tMax(4) = -1 And tMax(0) > tMax(3)) Then
                                tMax(2) = 3
                                tCoin(1, 0) = tMax(0)
                                tCoin(1, 1) = tMax(1)
                            End If
                        End If
                    End If
                End If
                vCouleur = ""
                tCount(1) = 1
            End If
        Next
        Close #FICHIERIMAGE
        ' Fin de la lecture

        FrmEnCours.Hide
        FrmMain.Show

        tCoin(3, 0) = Sqr((tCoin(0, 0) - tCoin(1, 0)) ^ 2 + (tCoin(0, 1) - tCoin(1, 1)) ^ 2)    ' Longueur d'un côté du rectangle
        tCoin(3, 1) = Sqr((tCoin(2, 0) - tCoin(1, 0)) ^ 2 + (tCoin(2, 1) - tCoin(1, 1)) ^ 2)    ' Longueur du 2ème côté du rectangle

        If ((tCercle(1) - tCercle(0)) / 2) ^ 2 * 3.141 > tPixNoir(2) - (tPixNoir(2) / 10) - (FileLen(vFichier) / vLong) And ((tCercle(1) - tCercle(0)) / 2) ^ 2 * 3.141 < tPixNoir(2) + (tPixNoir(2) / 10) + (FileLen(vFichier) / vLong) Then
            MsgBox "C'est un cercle", vbInformation, "Cercle"
        ElseIf tCoin(3, 0) * tCoin(3, 1) > tPixNoir(2) - (tPixNoir(2) / 10) And tCoin(3, 0) * tCoin(3, 1) < tPixNoir(2) + (tPixNoir(2) / 10) Then
            MsgBox "C'est un rectangle", vbInformation, "Rectangle"
        Else
            MsgBox "Aucuns résultats trouvés. Vérifiez que l'image n'est pas trop petite.", vbInformation, "No match found"
        End If
    End If
End Function

Private Sub CmdRefresh_Click()
    File1.Refresh
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim vTest As Boolean
    On Error GoTo NoDrive
    Dir1.Path = Drive1.Drive
    Exit Sub
NoDrive:
    If MsgBox("Le périphérique " & Drive1.Drive & " n'est pas disponible. Voulez-vous réessayer?", vbCritical + vbYesNo, "Erreur") = vbYes Then
        Drive1_Change
    Else
        Drive1.Drive = "c:"
    End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCreer_Click
    End If
End Sub

Private Sub File1_Click()
    File1.ToolTipText = File1.FileName
End Sub

Private Sub File1_DblClick()
    CmdOk_Click
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        MenuDel_Click
    End If
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If File1.FileName = "" Then
            MsgBox "Veuillez d'abord séléctionner un fichier", vbCritical, "Erreur"
        Else
            PopupMenu Menu
        End If
    End If
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:"
'    Drive1.Drive = "d:"
'    Dir1.Path = "\Faux TPI\Analyse d'image\Images"
    FrmMain.Top = 4000
    FrmMain.Left = 4000
End Sub

Private Sub MenuCopier_Click()
    vNomFichier = File1.Path & "\" & File1.FileName & Len(File1.Path)
    MenuColler.Enabled = True
End Sub

Private Sub MenuCouper_Click()
    If Len(File1.Path) < 10 Then
        vNomFichier = File1.Path & "\" & File1.FileName & "0" & Len(File1.Path) & "kill"
    Else
        vNomFichier = File1.Path & "\" & File1.FileName & Len(File1.Path) & "kill"
    End If
    MenuColler.Enabled = True
End Sub

Private Sub MenuColler_Click()
Dim vLong As String
On Error GoTo Fin
    If Right(vNomFichier, 4) = "kill" Then
        vLong = Right(vNomFichier, 6)
        vLong = Left(vLong, 2)
        FileCopy Left(vNomFichier, Len(vNomFichier) - 6), File1.Path & Mid(vNomFichier, vLong + 1, Len(vNomFichier) - vLong - 6)
        Kill Left(vNomFichier, Len(vNomFichier) - 6)
        MenuColler.Enabled = False
    Else
        vLong = Right(vNomFichier, 2)
        FileCopy Left(vNomFichier, Len(vNomFichier) - 2), File1.Path & Mid(vNomFichier, vLong + 1, Len(vNomFichier) - vLong - 2)
    End If
Fin:
    File1.Refresh
End Sub

Private Sub MenuDel_Click()
    If MsgBox("Voulez-vous vraiment supprimer cette image ?!?", vbYesNo, "Suppression") = vbYes Then
        Kill File1.Path & "\" & File1.FileName
        File1.Refresh
    End If
End Sub
