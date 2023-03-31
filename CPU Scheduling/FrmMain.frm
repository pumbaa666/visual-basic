VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPU Scheduling"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRAZ 
      Caption         =   "&Réinitialiser"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton CmdGen 
      Caption         =   "&Generer"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox TxtQuantum 
      Height          =   285
      Left            =   6480
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox ComboAlgorithmes 
      Height          =   315
      ItemData        =   "FrmMain.frx":0000
      Left            =   5400
      List            =   "FrmMain.frx":0010
      TabIndex        =   9
      Text            =   "Algorithmes"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton CmdCalculs 
      Caption         =   "Faire les &Calculs"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Ajouter &Entrée"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox LstDuree 
      Height          =   3765
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox LstArrivee 
      Height          =   3765
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox LstNom 
      Height          =   3765
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label LblQuantums 
      Caption         =   "Quantums :"
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Durée"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Arrivée"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nom"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAdd_Click()
    FrmAdd.Show
End Sub

Private Sub CmdGen_Click()
Dim vRand As Integer
Dim vCnt As Integer
Dim vCount As Integer
Dim vAdd As structEntree

    If LstNom.ListCount <> 0 Then
        ReDim tEntree(0)
        LstNom.Clear
        LstArrivee.Clear
        LstDuree.Clear
        vNbEntree = 0
        CmdRAZ_Click
    End If

    Randomize
    vRand = Int(Rnd * 6) + 3
    vNbEntree = 0

    For vCnt = 0 To vRand
        vAdd.vNom = "P" & vCnt
        vAdd.vArrivee = Int(Rnd * 70)
        vAdd.vDuree = Int(Rnd * 20) + 1

        LstNom.AddItem vAdd.vNom
        LstArrivee.AddItem vAdd.vArrivee
        LstDuree.AddItem vAdd.vDuree

        ReDim Preserve tEntree(vNbEntree)
        tEntree(vNbEntree).vNom = vAdd.vNom
        tEntree(vNbEntree).vArrivee = vAdd.vArrivee
        tEntree(vNbEntree).vDuree = vAdd.vDuree

        vNbEntree = vNbEntree + 1
    Next
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdCalculs_Click()
    Decharger
    If LstNom.ListCount <> 0 And ComboAlgorithmes.ListIndex <> -1 Then
        If ComboAlgorithmes.ListIndex = 3 And TxtQuantum.Text = "" Then
            MsgBox "Entrez le nombre de Quantum", vbCritical, "Erreur"
        Else
            If ComboAlgorithmes.ListIndex = 0 Then
                FCFS
                Chronogramme
            ElseIf ComboAlgorithmes.ListIndex = 1 Then
                SJFNot
                Chronogramme
            End If
        End If
    Else
        MsgBox "Il manque des paramètres", vbCritical, "Erreur"
    End If
End Sub

Private Sub ComboAlgorithmes_Click()
    If ComboAlgorithmes.ListIndex = 3 Then
        TxtQuantum.Visible = True
        LblQuantums.Visible = True
    Else
        LblQuantums.Visible = False
        TxtQuantum.Visible = False
    End If
End Sub

Private Sub CmdRAZ_Click()
    Decharger
    FrmListe.Liste(0).Clear
    FrmChrono.Hide
    FrmListe.Hide
    LstNom.Clear
    LstArrivee.Clear
    LstDuree.Clear
End Sub

Private Sub LstArrivee_Click()
    LstNom.ListIndex = LstArrivee.ListIndex
    LstDuree.ListIndex = LstArrivee.ListIndex
End Sub

Private Sub LstDuree_Click()
    LstArrivee.ListIndex = LstDuree.ListIndex
    LstNom.ListIndex = LstDuree.ListIndex
End Sub

Private Sub LstNom_Click()
    LstArrivee.ListIndex = LstNom.ListIndex
    LstDuree.ListIndex = LstArrivee.ListIndex
End Sub

Private Sub TxtQuantum_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        MsgBox "N'entrez que des chiffres", vbCritical, "Erreur"
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        CmdCalculs_Click
    End If
End Sub
