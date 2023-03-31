VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Economat"
   ClientHeight    =   6120
   ClientLeft      =   1620
   ClientTop       =   690
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListeQuantite 
      Height          =   4155
      Left            =   11880
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox ListeIngredients 
      Height          =   4155
      Left            =   9600
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CheckBox CheckTxt 
      Caption         =   "Voir uniquement les .txt"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton CmdBon 
      Caption         =   "&Faire le bon"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton CmdSel 
      Caption         =   "&Sélectionner"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   5400
      Width           =   2775
   End
   Begin VB.ListBox ListeRecette 
      Height          =   4155
      Left            =   6720
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   3480
      MultiSelect     =   2  'Extended
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Quantité :"
      Height          =   255
      Left            =   11880
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Ingrédients :"
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Liste des recettes sélectionnées :"
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Sélectionnez vos recettes :"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu MenuPopup 
      Caption         =   "MenuPopup"
      Visible         =   0   'False
      Begin VB.Menu PopupSupprimer 
         Caption         =   "Supprimer"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tListe(500, 1) As String
Dim vNbElem As Integer

Private Sub CheckTxt_Click()
    If CheckTxt.Value = 1 Then
        File1.Pattern = "*.txt"
    Else
        File1.Pattern = "*.*"
    End If
End Sub

Private Sub CmdBon_Click()
Dim i As Integer
Dim j As Integer
Dim vChemin As String
Dim vString As String
Dim vSplit() As String

    If ListeRecette.ListCount = 0 Then
        MsgBox "Vous n'avez choisi aucune recettes", vbCritical, "Connard"
    Else
        ViderTableau
        For i = 0 To ListeRecette.ListCount - 1
            ListeRecette.ListIndex = i
            vChemin = Dir1.Path & "\" & ListeRecette.Text

            Open vChemin For Input As #1
            Do While (Not EOF(1))
                Line Input #1, vString
                vSplit = Dissocier(vString, ",")
                Ajouter vSplit(0), vSplit(1)
                j = j + 1
                vNbElem = j
            Loop
            Close 1
        Next
    End If
    AfficherTot
End Sub

Private Function ViderTableau()
Dim i As Integer

    For i = 0 To vNbElem
        tListe(i, 0) = ""
        tListe(i, 1) = ""
    Next
    vNbElem = 0
End Function

Private Function AfficherTot()
Dim i As Integer

    ListeIngredients.Clear
    ListeQuantite.Clear
    For i = 0 To vNbElem
        If tListe(i, 0) <> "" Then
            ListeIngredients.AddItem tListe(i, 0)
            ListeQuantite.AddItem tListe(i, 1)
        End If
    Next
End Function

Private Function Ajouter(ByVal vIngredient As String, ByVal vQuantite As String)
Dim i As Integer
Dim vNew As Boolean

    vNew = True
    For i = 0 To vNbElem
        If tListe(i, 0) = vIngredient Then
            If Right(tListe(i, 1), 1) = "g" Then
                tListe(i, 1) = Int(Left(tListe(i, 1), Len(tListe(i, 1)) - 1)) + Int(Left(vQuantite, Len(vQuantite) - 1)) & "g"
            Else
                tListe(i, 1) = Int(tListe(i, 1)) + Int(vQuantite)
            End If
            vNew = False
            Exit For
        End If
    Next
    If vNew = True Then
        tListe(vNbElem, 0) = vIngredient
        tListe(vNbElem, 1) = vQuantite
    End If
End Function


Private Sub CmdSel_Click()
Dim i As Integer
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            If Not Existe(File1.List(i)) Then
                ListeRecette.AddItem (File1.List(i))
            End If
        End If
    Next
End Sub

Private Sub CmdQuitter_Click()
    End
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

Private Sub File1_DblClick()
    CmdSel_Click
End Sub

Private Function Existe(ByVal vNom As String)
Dim i As Integer
Dim vExiste As Boolean
    vExiste = False
    For i = 0 To ListeRecette.ListCount
        If ListeRecette.List(i) = vNom Then
            Existe = True
        End If
    Next
End Function

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CmdSel_Click
    End If
End Sub

Function Dissocier(Chaîne As String, Optional Balise As String = " ") As Variant
Dim Elements() As String
Dim LongChaîne As Long
Dim PrécBalise As Long
Dim K As Integer
PrécBalise = -1

    While (PrécBalise <> LongChaîne)
        LongChaîne = Len(Chaîne)
        Chaîne = Right(Chaîne, LongChaîne - PrécBalise - Len(Balise))
        'Recherche le dernier caractère avant la balise
        PrécBalise = InStr(1, Chaîne, Balise, vbBinaryCompare) - 1
        If PrécBalise = -1 Then PrécBalise = LongChaîne
        ReDim Preserve Elements(K)
        Elements(K) = Mid(Chaîne, 1, PrécBalise)
        K = K + 1
    Wend
    Dissocier = Elements()
End Function

Private Sub ListeRecette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MenuPopup
    End If
End Sub

Private Sub PopupSupprimer_Click()
Dim i As Integer

    For i = ListeRecette.ListCount - 1 To 0 Step -1
        ListeRecette.ListIndex = i
        If ListeRecette.Selected(i) Then
            ListeRecette.RemoveItem (i)
        End If
    Next
End Sub
