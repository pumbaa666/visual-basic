VERSION 5.00
Begin VB.Form FrmChercher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mes DVD - Chercher un DVD"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "&Rafficher toute la liste"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox TxtChercher 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtChercher 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxtChercher 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox TxtChercher 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox TxtChercher 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.ComboBox ComboNote 
      Height          =   315
      ItemData        =   "FrmChercher.frx":0000
      Left            =   1080
      List            =   "FrmChercher.frx":0002
      TabIndex        =   3
      Text            =   "1 --> 9"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CmdChercher 
      Caption         =   "&Chercher"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Titre :"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Genre :"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Acteur(s) :"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Note :"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label LblPreter 
      Alignment       =   1  'Right Justify
      Caption         =   "Prêté à :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
End
Attribute VB_Name = "FrmChercher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNumChercher As Integer

Private Sub CmdAnnuler_Click()
    CmdRefresh_Click
    FrmChercher.Hide
End Sub

Private Sub CmdChercher_Click()
Dim vLigne As Integer
Dim vColonne As Integer
Dim vNb As Integer

    If TxtChercher(0).Text = "" And TxtChercher(1).Text = "" And TxtChercher(2).Text = "" And (ComboNote.Text = "" Or ComboNote.Text = "1 --> 9") And TxtChercher(4).Text = "" And vNumChercher <> 4 Then
        MsgBox "Veuillez entrer au moins un paramètre.", vbCritical, "Erreur"
    Else
'        If vNumChercher = 4 And TxtChercher(4) = "" Then
'            vchercher = ""
'        End If
        ClearListe
        For vLigne = 1 To vNbDVDTot '- 1
            '*** faire les tests les uns après les autres (titre, genre...) sans elseif ***'
            '*** changer le vnumchercher ***'
            If ComboNote.Text = "Sans notes" And tListe(4, vLigne) = "" Then
                FrmListe.Liste(0).AddItem vLigne
                vNb = vNb + 1
                For vColonne = 1 To 5
                    FrmListe.Liste(vColonne).AddItem tListe(vColonne, vLigne)
                Next
            ElseIf InStr(1, LCase(tListe(vNumChercher + 1, vLigne)), LCase(TxtChercher(vNumChercher).Text), vbTextCompare) <> 0 Then
                FrmListe.Liste(0).AddItem vLigne
                vNb = vNb + 1
                For vColonne = 1 To 5
                    FrmListe.Liste(vColonne).AddItem tListe(vColonne, vLigne)
                Next
            End If
        Next
        For vColonne = 0 To 5
            FrmListe.Liste(vColonne).AddItem ""
        Next
        FrmListe.Liste(0).AddItem "Total : "
        FrmListe.Liste(1).AddItem vNb
        For vColonne = 2 To 5
            FrmListe.Liste(vColonne).AddItem ""
        Next
    End If
End Sub

Private Sub CmdRefresh_Click()
    vModif = False
    RefreshListe
End Sub

Private Sub ComboNote_Click()
    TxtChercher_GotFocus 5
    TxtChercher(3).Text = ComboNote.Text
End Sub

Private Sub Form_Activate()
Dim vCount As Integer

    If vRecherche = 0 Then
        TxtChercher_GotFocus 5
        ComboNote.Clear
        ComboNote.AddItem "Sans notes"
        For vCount = 0 To 9
            ComboNote.AddItem vCount
        Next
        ComboNote.Text = "1 --> 9"
    End If
End Sub

Private Sub Form_Load()
Dim vCount As Integer
    ComboNote.Clear
    ComboNote.AddItem "Sans notes"
    For vCount = 0 To 9
        ComboNote.AddItem vCount
    Next
End Sub

Private Sub TxtChercher_GotFocus(Index As Integer)
Dim vCount As Integer

    If Index = 5 Then
       vNumChercher = 3
    Else
        vNumChercher = Index
    End If
    For vCount = 0 To 4
        TxtChercher(vCount) = ""
    Next
End Sub

Private Sub TxtChercher_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdChercher_Click
    End If
End Sub
