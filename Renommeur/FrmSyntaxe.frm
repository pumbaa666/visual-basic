VERSION 5.00
Begin VB.Form FrmSyntaxe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comment renommer"
   ClientHeight    =   7350
   ClientLeft      =   345
   ClientTop       =   390
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Majuscules"
      Height          =   1695
      Left            =   240
      TabIndex        =   21
      Top             =   5400
      Width           =   5295
      Begin VB.OptionButton OptNulPart 
         Caption         =   "Nul part"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton OptPartout 
         Caption         =   "Partout"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton OptChaque 
         Caption         =   "Au début de chaque mot"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton OptPremier 
         Caption         =   "1er mot seulement"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.CommandButton CmdOkMaj 
         Caption         =   "Ok"
         Height          =   735
         Left            =   4080
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Renommer partiellement"
      Height          =   2055
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   5295
      Begin VB.CheckBox ChkCasse 
         Caption         =   "Respecter la casse (Case sensitive)"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Fait la distinction entre majuscule et minuscule"
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.OptionButton OptUn 
         Caption         =   "Remplacer une fois "
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         ToolTipText     =   "Ne remplace que la 1ère chaine trouvée"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton OptTout 
         Caption         =   "Remplacer tout"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Remplace toute les occurences"
         Top             =   1320
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton CmdOkPart 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TxtRemp 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Text            =   "Caractère ou chaine"
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox TxtCherche 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Text            =   "Caractère ou chaine"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Remplacer par : "
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Rechercher : "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Renommer totalement"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   1800
         Width           =   2415
      End
      Begin VB.OptionButton OptNumText 
         Caption         =   "N° + Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptTextNum 
         Caption         =   "Text + N°"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "N° d'incrémentation"
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2175
         Begin VB.TextBox TxtStart 
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "0"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TxtInc 
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "1"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Partir de"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Incrémenter de"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtIns 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Text            =   "Insérez votre texte ici (facultatif)"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox ChkMulti 
         Caption         =   "&Activer le ""multi-zéro"""
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         ToolTipText     =   "Place des 0 devant les chiffres pour faire en sorte que les fichiers soient bien triés par ordre alphabetique"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Text"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmSyntaxe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
Dim vCount As Integer
Dim vCheminFichier As String
Dim vAffCount As String
Static vInserer As String
Dim vNbZero As Integer
Dim vCntZero As Integer
Dim vZero As String

    On Error GoTo AlreadyExist
    If FrmMain.File1.ListCount = 0 Then
        MsgBox "Il n'y a pas de fichiers", vbCritical, "Erreur"
    ElseIf TxtStart.Text = "" Or TxtInc.Text = "" Then
        MsgBox "Veuillez compléter les champs N° d'incrémentations", vbCritical, "Erreur"
    Else
        FrmMain.File1.Refresh

        If Len(FrmMain.Dir1.Path) = 3 Then
            vCheminFichier = FrmMain.Dir1.Path
        Else
            vCheminFichier = FrmMain.Dir1.Path & "\"
        End If

        For vCount = 0 To FrmMain.File1.ListCount - 1
            FrmMain.File1.ListIndex = vCount    ' Sélectionne le fichier suivant
            vNbZero = Len(Trim(Str(Int(FrmMain.File1.ListCount) * Int(TxtInc.Text) + Int(TxtStart.Text)))) - Len(Trim(Str(Int(vCount) * Int(TxtInc.Text) + Int(TxtStart.Text))))
            vZero = ""

            If ChkMulti <> Checked Then
                Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & TxtIns.Text & vCount * TxtInc.Text + TxtStart.Text & LCase(Right(FrmMain.File1.FileName, 4))
            Else
                For vCntZero = 0 To vNbZero - 1
                    vZero = vZero & "0"
                Next

                If OptTextNum.Value = True Then
                    Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & TxtIns.Text & vZero & (vCount * TxtInc.Text + TxtStart.Text) & LCase(Right(FrmMain.File1.FileName, 4))
                Else
                    Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & vZero & (vCount * TxtInc.Text + TxtStart.Text) & TxtIns.Text & LCase(Right(FrmMain.File1.FileName, 4))
                End If
            End If
        Next
        FrmMain.File1.Refresh
        If vInserer = "" Then   ' N'affiche pas le message quand je fais un 1er renommage bidon
            If FrmMain.File1.ListCount = 1 Then
                MsgBox "Action terminée, 1 fichier renommé", vbInformation, "Terminé"
            Else
                MsgBox "Action terminée, " & FrmMain.File1.ListCount & " fichiers renommés", vbInformation, "Terminé"
            End If
        End If
    End If
    Exit Sub

AlreadyExist:
    vInserer = TxtIns.Text
    TxtIns.Text = "abcdefg#@"
    CmdOk_Click
    TxtIns.Text = vInserer
    vInserer = ""
    CmdOk_Click
End Sub

Private Sub CmdOkMaj_Click()
Dim vCount As Integer
Dim vCount2 As Integer
Dim vCheminFichier As String
Dim vAffCount As String
Static vInserer As String
Dim vNewName As String
Dim vPosCar As Long
Dim vNbFichier As Integer
Dim vCar As String

    If FrmMain.File1.ListCount = 0 Then
        MsgBox "Il n'y a pas de fichiers", vbCritical, "Erreur"
    Else
        FrmMain.File1.Refresh

        If Len(FrmMain.Dir1.Path) = 3 Then
            vCheminFichier = FrmMain.Dir1.Path
        Else
            vCheminFichier = FrmMain.Dir1.Path & "\"
        End If

        vNbFichier = 0
        For vCount = 0 To FrmMain.File1.ListCount - 1
            FrmMain.File1.ListIndex = vCount    ' Sélectionne le fichier suivant
    
            If OptPremier.Value = True Then
                vNewName = UCase(Left(FrmMain.File1.FileName, 1)) & Right(FrmMain.File1.FileName, Len(FrmMain.File1.FileName) - 1)
                Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & vNewName
            ElseIf OptChaque.Value = True Then
                vNewName = UCase(Left(FrmMain.File1.FileName, 1)) ' 1ère lettre
                vCount2 = 2
                Do
                    If Mid(FrmMain.File1.FileName, vCount2 - 1, 1) = " " Then
                        vCar = UCase(Mid(FrmMain.File1.FileName, vCount2, 1))
                    Else
                        vCar = Mid(FrmMain.File1.FileName, vCount2, 1)
                    End If
                    vNewName = vNewName & vCar
                    vCount2 = vCount2 + 1
                Loop While (vCount2 <= Len(FrmMain.File1.FileName))
                Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & vNewName
            ElseIf OptPartout.Value = True Then
                vNewName = UCase(FrmMain.File1.FileName)
                Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & vNewName
            Else
                vNewName = LCase(FrmMain.File1.FileName)
                Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & vNewName
            End If
            vNbFichier = vNbFichier + 1
        Next

        FrmMain.File1.Refresh
        If vNbFichier = 0 Then
            MsgBox "Aucun fichier n'a été renommé", vbInformation, "Terminé"
        ElseIf vNbFichier = 1 Then
            MsgBox "Action terminée, 1 fichier renommé", vbInformation, "Terminé"
        Else
            MsgBox "Action terminée, " & vNbFichier & " fichiers renommés", vbInformation, "Terminé"
        End If
    End If
End Sub

Private Sub CmdOkPart_Click()
Dim vCount As Integer
Dim vCheminFichier As String
Dim vAffCount As String
Static vInserer As String
Dim vNewName As String
Dim vPosCar As Long
Dim vNbFichier As Integer
Dim vSauvFichier As Integer

    If FrmMain.File1.ListCount = 0 Then
        MsgBox "Il n'y a pas de fichiers", vbCritical, "Erreur"
    ElseIf TxtCherche.Text = "" Or TxtCherche.Text = "Caractère ou chaine" Then
        MsgBox "Veuillez indiquer se que vous voulez chercher", vbCritical, "Erreur"
    Else
        Do
            FrmMain.File1.Refresh

            If Len(FrmMain.Dir1.Path) = 3 Then
                vCheminFichier = FrmMain.Dir1.Path
            Else
                vCheminFichier = FrmMain.Dir1.Path & "\"
            End If

            vNbFichier = 0
            For vCount = 0 To FrmMain.File1.ListCount - 1
                FrmMain.File1.ListIndex = vCount    ' Sélectionne le fichier suivant
    
                If ChkCasse.Value = Checked Then
                    vPosCar = InStr(1, FrmMain.File1.FileName, TxtCherche.Text, vbTextCompare)
                    If vPosCar <> 0 Then
                        If Mid(FrmMain.File1.FileName, vPosCar, Len(TxtCherche.Text)) <> TxtCherche.Text Then
                            vPosCar = 0
                        End If
                    End If
                Else
                    vPosCar = InStr(1, FrmMain.File1.FileName, TxtCherche.Text, vbTextCompare)
                End If

                If vPosCar <> 0 Then
                    vNewName = Left(FrmMain.File1.FileName, vPosCar - 1) & TxtRemp.Text & Right(FrmMain.File1.FileName, Len(FrmMain.File1.FileName) - vPosCar - Len(TxtCherche.Text) + 1)
                    Name vCheminFichier & FrmMain.File1.FileName As vCheminFichier & vNewName
                    vNbFichier = vNbFichier + 1
                    vSauvFichier = vSauvFichier + 1
                End If
            Next
            If OptUn.Value = True Then 'Or InStr(1, TxtRemp.Text, TxtCherche.Text, vbTextCompare) <> 0 Then
                vNbFichier = 0
            ElseIf vCount = 0 Then
                vSauvFichier = vNbFichier
            End If
        Loop While (vNbFichier <> 0)

        FrmMain.File1.Refresh
        If vSauvFichier = 0 Then
            MsgBox "Aucun fichier n'a été renommé", vbInformation, "Terminé"
        ElseIf vSauvFichier = 1 Then
            MsgBox "Action terminée, 1 fichier renommé", vbInformation, "Terminé"
        Else
            MsgBox "Action terminée, " & vSauvFichier & " fichiers renommés", vbInformation, "Terminé"
        End If
    End If
End Sub

Private Sub TxtIns_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vPos As Long

    vPos = InStrRev(FrmMain.Dir1.Path, "\", Len(FrmMain.Dir1.Path) - 1, vbTextCompare)
    TxtIns.Text = Right(FrmMain.Dir1.Path, Len(FrmMain.Dir1.Path) - vPos) & " "
End Sub

Private Sub TxtRemp_GotFocus()
    If TxtRemp.Text = "Caractère ou chaine" Then
        TxtRemp.Text = ""
    End If
End Sub

Private Sub TxtCherche_GotFocus()
    If TxtCherche.Text = "Caractère ou chaine" Then
        TxtCherche.Text = ""
    End If
End Sub

Private Sub TxtInc_KeyPress(KeyAscii As Integer)
    If (KeyAscii <= 48 Or KeyAscii >= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtIns_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtRemp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOkPart_Click
    End If
End Sub

Private Sub TxtStart_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
