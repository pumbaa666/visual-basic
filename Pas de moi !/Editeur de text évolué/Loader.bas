Attribute VB_Name = "Loader"

'#######################################################
'#MODULE DE CHARGEMENT ET D'ENREGISTREMENT DES FICHIERS#
'#######################################################

Public FichierActuel As String
Public MustBeSaved As Boolean
Public AlreadySaved As Boolean

Public Sub LoadFileZ(Fichier As String)

Dim LastStatus As String

'Fonction qui va analyser le fichier demandé et l'envoyer
'à la bonne fonction de chargement

AlreadySaved = True
MustBeSaved = False
LockMain
LastStatus = Main.Label.Caption
Status "Ouverture en cours..."

If FileLen(Fichier) > 0 Then
    Select Case LCase(Right(Fichier, 3))
        Case "txt", "log", "rtx", "wtx"
            Main.Text.Text = LoadTXT(Fichier)
        Case "rtf"
            LoadRTF Fichier
        Case Else
            If MessageBox.Message("Ce type de fichier n'est pas pris en charge." & vbCrLf & "Voulez-vous qu'il soit ouvert en mode texte?", "Erreur lors de l'ouverture", YesNo, Information, Main) = Yes Then
                Main.Text.Text = LoadTXT(Fichier)
            Else
                Status LastStatus
                UnlockMain
                Exit Sub
            End If
    End Select
End If

TamponTexte = Main.Text.Text
Status FichierActuel
UnlockMain

End Sub

Private Function LoadTXT(Fichier As String) As String

Dim Buffer, Buffer2 As String

'Fonction de chargement de fichiers texte de base (.txt,.log...)

With Main
    .Progress.Maxi = FileLen(Fichier)
    Open Fichier For Input As #1
        Do While Not EOF(1)
            DoEvents
            Line Input #1, Buffer
            Buffer2 = Buffer2 & Buffer & vbCrLf
            If EOF(1) Then Exit Do
            If Len(Buffer2) > 10000 Then .Progress.Value = .Progress.Value + Len(Buffer2): LoadTXT = LoadTXT & Buffer2: Buffer2 = ""
        Loop
    Close #1
    If Len(Buffer2) <> 0 Then LoadTXT = LoadTXT & Left(Buffer2, Len(Buffer2) - 2)
    .Progress.Value = 0
End With

End Function

Private Function LoadRTF(Fichier As String) As String

'Fonction de chargement de fichiers RTF
Main.Text.LoadFile Fichier

End Function

Public Sub SaveFileZ(Texte As String, Fichier As String)

Dim LastStatus As String

'Fonction qui va analyser le fichier demandé et l'envoyer
'à la bonne fonction d'enregistrement

AlreadySaved = True
MustBeSaved = False
LockMain
LastStatus = Main.Label.Caption
Status "Enregistrement en cours..."

Select Case LCase(Right(Fichier, 3))
    Case "txt", "log", "rtx", "wtx"
        SaveTXT Texte, Fichier
    Case "rtf"
        SaveRTF Texte, Fichier
    Case Else
        MessageBox.Message "Ce type de fichier n'est pas pris en charge.", "Erreur lors de l'enregistrement", OkOnly, Information, Main
        Status LastStatus
        UnlockMain
        Exit Sub
End Select

TamponTexte = Main.Text.Text
UnlockMain
Status FichierActuel

End Sub

Private Sub SaveTXT(Texte As String, Fichier As String)

'Fonction de chargement de fichiers texte de base (.txt,.log...)

Open Fichier For Output As #1
    Print #1, Texte
Close #1

End Sub

Private Sub SaveRTF(Texte As String, Fichier As String)

'Fonction de chargement de fichiers RTF
Main.Text.SaveFile Fichier

End Sub



