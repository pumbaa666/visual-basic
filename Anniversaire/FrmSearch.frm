VERSION 5.00
Begin VB.Form FrmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rechercher"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Prochain anniversaire"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ListBox ListTrouve 
      Height          =   3375
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   5175
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Chercher"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxtSearch 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.OptionButton OptJour 
      Caption         =   "Jour"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton OptMois 
      Caption         =   "Mois"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.OptionButton OptAnnee 
      Caption         =   "Année"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton OptNom 
      Caption         =   "Nom"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptPrenom 
      Caption         =   "Prénom"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton OptMail 
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Visible         =   0   'False
      Begin VB.Menu MenuFichierDel 
         Caption         =   "Supprimer"
      End
      Begin VB.Menu MenuFichierMail 
         Caption         =   "Envoyer un mail"
      End
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tPlace(50) As Integer
Dim vEndroit As Integer
Dim vAt As Integer
Dim vMail As String

Private Sub CmdAnnuler_Click()
    FrmMain.Show
    FrmSearch.Hide
End Sub

Private Sub CmdNext_Click()
Dim vNext As Date
Dim vCount As Integer
Dim vEnCours As String
Dim tTout(3) As String
Dim tToutEnCours(3) As String
Dim vLast As String
Dim vMois As String
Dim vNbJours As Integer

    Open "c:\temp\donnees.dat" For Input As #1
    Do
        Line Input #1, vEnCours
        tToutEnCours(vCount) = vEnCours
        If vLast = tTout(2) And vCount = 3 Then
            tTout(3) = vEnCours
        End If
        vLast = vEnCours
        If vCount = 2 Then
            If vNext = "00:00:00" Then
                If CDate(Left(vEnCours, 5)) > CDate(Left(Now, 5)) Then
                    vNext = vEnCours
                    tTout(0) = tToutEnCours(0)
                    tTout(1) = tToutEnCours(1)
                    tTout(2) = tToutEnCours(2)
                End If
            Else
                If CDate(Left(vEnCours, 5)) > CDate(Left(Now, 5)) And CDate(Left(vEnCours, 5)) < CDate(Left(vNext, 5)) Then
                    vNext = vEnCours
                End If
            End If
        End If
        vCount = vCount + 1
        If vCount = 4 Then vCount = 0
    Loop Until (EOF(1))
    Close #1

    ListTrouve.Clear
    If vNext = "00:00:00" Then
        ListTrouve.AddItem "Il n'y aura plus d'anniversaire cette année"
    Else
        vMois = Switch(Month(vNext) = 1, "Janvier", Month(vNext) = 2, "Février", _
                       Month(vNext) = 3, "Mars", Month(vNext) = 4, "Avril", _
                       Month(vNext) = 5, "Mai", Month(vNext) = 6, "Juin", _
                       Month(vNext) = 7, "Juillet", Month(vNext) = 8, "Août", _
                       Month(vNext) = 9, "Septembre", Month(vNext) = 10, "Octobre", _
                       Month(vNext) = 11, "Novembre", Month(vNext) = 12, "Décembre")
        ListTrouve.AddItem "Le prochain anniversaire aura lieu le " & Day(vNext) & " " & vMois
        ListTrouve.AddItem ""
        ListTrouve.AddItem tTout(0) & " " & tTout(1) & " " & tTout(2) & " " & tTout(3)
        ListTrouve.AddItem ""
        ListTrouve.AddItem ""
    End If
End Sub

Private Sub CmdSearch_Click()
Dim vParam As Integer
Dim tDonnee(3) As String
Dim vCount As Integer
Dim vCountGen As Integer
Dim vCountGen2 As Integer
Dim vNbTrouve As Integer
Dim vCherche As String
    
' Protection de saisie
    If TxtSearch.Text = "" Then
        MsgBox "Entrez un critère", vbCritical, "Erreur"
        Exit Sub
    ElseIf OptJour.Value = True Then
        If Len(TxtSearch.Text) <> 2 Or Int(TxtSearch.Text) > 31 Then
            MsgBox "Saisie incorrecte", vbCritical, "Erreur"
            Exit Sub
        End If
    ElseIf OptMois.Value = True Then
        If Len(TxtSearch.Text) <> 2 Or Int(TxtSearch.Text) > 12 Then
            MsgBox "Saisie incorrecte", vbCritical, "Erreur"
            Exit Sub
        End If
    ElseIf OptAnnee.Value = True Then
        If Len(TxtSearch.Text) <> 4 Or Int(TxtSearch.Text) > Year(Date) Then
            MsgBox "Saisie incorrecte", vbCritical, "Erreur"
            Exit Sub
        End If
    End If
    
' Recherche
    ListTrouve.Clear
    ListTrouve.AddItem "Résultat de la recherche pour " & Trim(TxtSearch)
    ListTrouve.AddItem ""
    Open "c:\temp\donnees.dat" For Input As #1
    If OptNom.Value = True Then
        vParam = 0
    ElseIf OptPrenom.Value = True Then
        vParam = 1
    ElseIf OptMail.Value = True Then
        vParam = 3
    ElseIf OptJour.Value = True Then
        vParam = 4
    ElseIf OptMois.Value = True Then
        vParam = 5
    ElseIf OptAnnee.Value = True Then
        vParam = 6
    End If
    Do
        On Error Resume Next
        For vCount = 0 To 3
            Line Input #1, tDonnee(vCount)
        Next
        
        If vParam = 4 Then
            vCherche = Left(tDonnee(2), 2)
        ElseIf vParam = 5 Then
            vCherche = Mid(tDonnee(2), 4, 2)
        ElseIf vParam = 6 Then
            vCherche = Right(tDonnee(2), 4)
        Else
            vCherche = LCase(tDonnee(vParam))
        End If
        
        If InStr(1, vCherche, Trim(LCase(TxtSearch.Text)), vbTextCompare) <> 0 Then
            ListTrouve.AddItem tDonnee(0) & " " & tDonnee(1) & " " & tDonnee(2) & " " & tDonnee(3)
            vNbTrouve = vNbTrouve + 1
            tPlace(vCountGen2) = vCountGen
            vCountGen2 = vCountGen2 + 1
        End If
        vCountGen = vCountGen + 1
    Loop Until (EOF(1))
    Close #1
    If vNbTrouve = 0 Then
        ListTrouve.AddItem "Aucuns résultats"
    Else
        ListTrouve.AddItem "________________________________________________________________________"
        ListTrouve.AddItem "Total : " & vNbTrouve
    End If
End Sub

Private Sub Form_Activate()
    ListTrouve.Clear
    ListTrouve.AddItem "Choisissez le mode de recherche"
    ListTrouve.AddItem "Entrez une valeur"
    ListTrouve.AddItem "Et cliquez sur Chercher"
End Sub

Private Sub ListTrouve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        If Int(Y / 195) > 1 And Int(Y / 195) < ListTrouve.ListCount - 2 Then
            ListTrouve.ListIndex = Int(Y / 195)
            vAt = InStr(1, ListTrouve.Text, "@", vbTextCompare)
            vMail = ListTrouve.Text
            If vAt = 0 Then
                MenuFichierMail.Enabled = False
            Else
                MenuFichierMail.Enabled = True
            End If
            PopupMenu MenuFichier
        End If
    ElseIf Button = 1 Then
        vEndroit = tPlace(ListTrouve.ListIndex - 2)
    End If
End Sub

Private Sub MenuFichierDel_Click()
    fSupprimer vEndroit
End Sub

Private Sub MenuFichierMail_Click()
Dim vAdresse As String

    Do
        If Mid(vMail, vAt - 1, 1) = " " Then
            vAdresse = Right(vMail, Len(vMail) - vAt + 1)
            vAt = 1
        End If
        vAt = vAt - 1
    Loop While (vAt <> 0)
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & vAdresse & "?subject=Anniversaire", vbMaximizedFocus
End Sub

Private Sub OptAnnee_Click()
    TxtSearch.Enabled = True
    TxtSearch.MaxLength = 4
    If Not IsNumeric(TxtSearch.Text) Then
        TxtSearch.Text = ""
    End If
End Sub

Private Sub OptJour_Click()
    TxtSearch.Enabled = True
    TxtSearch.MaxLength = 2
    If Not IsNumeric(TxtSearch.Text) Then
        TxtSearch.Text = ""
    End If
End Sub

Private Sub OptMail_Click()
    TxtSearch.Enabled = True
    TxtSearch.MaxLength = 0
End Sub

Private Sub OptMois_Click()
    TxtSearch.Enabled = True
    TxtSearch.MaxLength = 2
    If Not IsNumeric(TxtSearch.Text) Then
        TxtSearch.Text = ""
    End If
End Sub

Private Sub OptNext_Click()
    TxtSearch.Enabled = False
End Sub

Private Sub OptNom_Click()
    TxtSearch.Enabled = True
    TxtSearch.MaxLength = 0
End Sub

Private Sub OptPrenom_Click()
    TxtSearch.Enabled = True
    TxtSearch.MaxLength = 0
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    If OptJour.Value = True Or OptMois.Value = True Or OptAnnee.Value = True Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
            MsgBox "N'entrez que des chiffres", vbCritical, "Erreur"
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        CmdSearch_Click
    End If
End Sub
