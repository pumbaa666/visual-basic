VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorateur"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
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
      Width           =   2775
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
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNomFichier As String

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdOk_Click()
Dim tCount(2), tPixNoir(7), vLong As Double ' tPixNoir : 0=En cours  1=Max  2=Tot  3=CooX  4=CooY  5=CooX Old  6=CooY Old  7=PosXMax
Dim vCouleur, vChaineTot, vFichier As String
Dim vSaut, vHaut, tCote(1), tCercle(1) As Integer
Dim vTemp As Byte
tCount(1) = 0
tCount(2) = 0
    
    If File1.FileName = "" Then
        MsgBox "Veuillez choisire une image", vbCritical, "Erreur"
    Else
        If Len(Dir1.Path) = 3 Then
            vFichier = Dir1.Path & File1.FileName
        Else
            vFichier = Dir1.Path & "\" & File1.FileName
        End If
        
        Open vFichier For Random As #1 Len = 1
        Get #1, 19, vTemp
        vLong = vTemp
        Get #1, 20, vTemp
        vLong = vLong + 256 * vTemp
'        Close #1

        If vLong Mod 4 = 0 Then
            vSaut = 0
        ElseIf vLong Mod 3 = 0 Then
            vSaut = 3
        ElseIf vLong Mod 2 = 0 Then
            vSaut = 2
        Else
            vSaut = 1
        End If

'        List1.Clear

'        Open vFichier For Random As #1 Len = 1
'        For tCount(0) = FileLen(vFichier) To 54 Step -1
        For tCount(0) = 54 To FileLen(vFichier)
            If tCount(0) = 54 + vLong * (3 + tCount(2)) Then
                tCount(0) = tCount(0) + vSaut
                tCount(2) = tCount(2) + vSaut + (4 - vSaut)
            End If

            Get #1, tCount(0), vTemp
            vCouleur = vTemp & ", " & vCouleur
            tCount(1) = tCount(1) + 1
            If tCount(1) = 4 Then
                If Val(Left(vCouleur, 3)) < 75 And Val(Mid(vCouleur, 6, 3)) < 75 And Val(Mid(vCouleur, 11, 3)) < 75 Then   'Si il y a un pixel noir
                    
                    tPixNoir(0) = tPixNoir(0) + 1
                    tPixNoir(2) = tPixNoir(2) + 1

                    tPixNoir(5) = tPixNoir(3)
                    tPixNoir(6) = tPixNoir(4)

                    tPixNoir(4) = Int((((tCount(0) - 54) / 3) - 2) / vLong)
                    tPixNoir(3) = (tCount(0) - 54) / 3 - tPixNoir(4) * vLong
                    
                    If tCercle(0) = 0 Then
                        tCercle(0) = tPixNoir(4)
                    End If
                    If tPixNoir(4) > 50 Then
                        tCercle(1) = tPixNoir(4)
                    End If

'                    If FrmImage.ChkCote.Value = Checked Then
                        If tPixNoir(3) <> tPixNoir(5) And tPixNoir(4) <> tPixNoir(6) Then
                            
                            If tPixNoir(6) > tPixNoir(5) Then
                                tCote(0) = tCote(0) + 1
                            Else
                                tCote(1) = tCote(1) + 1
                            End If
                            
'                            List1.AddItem tPixNoir(3) & ", " & tPixNoir(4)
                            FrmImage.PSet (tPixNoir(3) * 15 + 500, tPixNoir(4) * 15 + 500)
'                        End If
'                    Else
'                        List1.AddItem tPixNoir(3) & ", " & tPixNoir(4)
'                        FrmImage.PSet (tPixNoir(3) * 15 + 500, tPixNoir(4) * 15 + 500)
                    End If
                End If
                vCouleur = ""
                tCount(1) = 1
            End If
        Next
        Close #1
'        FrmImage.LblX.Caption = tCercle(0)
'        FrmImage.LblY.Caption = tCercle(1)
'        FrmImage.Label1.Caption = "Tot trouvé : " & tPixNoir(2)
'        FrmImage.Label2.Caption = "Tot calculé : " & ((tCercle(1) - tCercle(0)) / 2) ^ 2 * 3.141
        MsgBox "L'analyse est finie, il y a " & Int(tPixNoir(2)) & " pixels noirs.", vbInformation, "Analyse"
        If 2 * (tCote(0) * tCote(1)) > tPixNoir(2) - 100 And 2 * (tCote(0) * tCote(1) < tPixNoir(2) + 100) Then
            MsgBox "C'est un rectangle"
        ElseIf ((tCercle(1) - tCercle(0)) / 2) ^ 2 * 3.141 > tPixNoir(2) - 100 And ((tCercle(1) - tCercle(0) / 2)) ^ 2 * 3.141 > tPixNoir(2) + 100 Then
            MsgBox "C'est un cercle"
        End If
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim vTest As Boolean
    On Error GoTo NoDrive
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    vTest = True
NoDrive:
If vTest = False Then
    MsgBox "Le périphérique " & Drive1.Drive & " n'est pas disponible.", vbCritical, "Erreur"
    Drive1.Drive = "c:"
End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCreer_Click
    End If
End Sub

Private Sub File1_DblClick()
    CmdOk_Click
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
    If File1.ListCount > 0 Then
        If KeyCode = 8 Or KeyCode = 46 Then
            If MsgBox("Voulez-vous vraiment supprimer cette image ?!?", vbYesNo, "Suppression") = vbYes Then
                Kill File1.Path & "\" & File1.FileName
                File1.Refresh
            End If
        End If
    Else
        MsgBox "Il n'y a pas d'image à supprimer!", vbCritical, "Erreur"
    End If
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Menu
    End If
End Sub

Private Sub Form_Load()
'    Drive1.Drive = "c:"
    Drive1.Drive = "d:"
    Dir1.Path = "\Faux TPI\Analyse d'image\Images"
    FrmImage.Show
    FrmMain.Top = 500
    FrmMain.Left = 500
    FrmImage.Top = 500
    FrmImage.Left = 7000
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
    File1.Refresh
Fin:
End Sub
