VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mes DVD - Ajouter un DVD"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtNum 
      Height          =   285
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox TxtPrete 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton CmdAnnuler 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox ComboNote 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Text            =   "1 --> 9"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TxtActeurs 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtGenre 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TxtTitre 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label LblPrete 
         Caption         =   "Prêté à :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   15
         Left            =   0
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Note :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Acteur(s) :"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Titre :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Numéro :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    FrmAdd.Hide
End Sub

Private Sub CmdOk_Click()
Dim vTempNbDVD As Integer

    If TxtTitre.Text = "" Or TxtGenre.Text = "" Then
        MsgBox "Vous devez au moins rentrer le titre et le genre du film", vbCritical, "Erreur"
    Else
        If Mid(FrmAdd.Caption, 11, 5) = "Modif" Then
            vTempNbDVD = vNbDVDTot
            vNbDVDTot = TxtNum.Text - 1
            Ajouter TxtTitre.Text, TxtGenre.Text, TxtActeurs.Text, ComboNote.Text, TxtPrete.Text
            vNbDVDTot = vTempNbDVD
            FrmListe.Liste(1).RemoveItem Int(TxtNum.Text) - 1
            FrmListe.Liste(1).AddItem TxtTitre.Text, Int(TxtNum.Text) - 1
            FrmListe.Liste(2).RemoveItem Int(TxtNum.Text) - 1
            FrmListe.Liste(2).AddItem TxtGenre.Text, Int(TxtNum.Text) - 1
            FrmListe.Liste(3).RemoveItem Int(TxtNum.Text) - 1
            FrmListe.Liste(3).AddItem TxtActeurs.Text, Int(TxtNum.Text) - 1
            If ComboNote.Text <> "1 --> 9" Then
                FrmListe.Liste(4).RemoveItem Int(TxtNum.Text) - 1
                FrmListe.Liste(4).AddItem ComboNote.Text, Int(TxtNum.Text) - 1
            End If
            FrmListe.Liste(5).RemoveItem Int(TxtNum.Text) - 1
            FrmListe.Liste(5).AddItem TxtPrete.Text, Int(TxtNum.Text) - 1
'            If Int(TxtNum.Text < vNbDVDTot) Then
                FrmListe.Liste(0).ListIndex = Int(TxtNum.Text) - 1
'            End If
        Else
            Ajouter TxtTitre.Text, TxtGenre.Text, TxtActeurs.Text, ComboNote.Text, ""
            FrmListe.Liste(0).AddItem vNbDVDTot
            FrmListe.Liste(1).AddItem TxtTitre.Text
            FrmListe.Liste(2).AddItem TxtGenre.Text
            FrmListe.Liste(3).AddItem TxtActeurs.Text
            If ComboNote.Text <> "1 --> 9" And ComboNote.Text <> "Sans notes" Then
                FrmListe.Liste(4).AddItem ComboNote.Text
            Else
                FrmListe.Liste(4).AddItem ""
            End If
            FrmListe.Liste(5).AddItem ""
            FrmListe.Liste(0).ListIndex = vNbDVDTot - 1
        End If
    End If
End Sub

Private Sub Form_Activate()
Dim vCount As Integer
    If Mid(FrmAdd.Caption, 11, 5) <> "Modif" Then
        TxtTitre.Text = ""
        TxtGenre.Text = ""
        TxtActeurs.Text = ""
    End If
    ComboNote.Clear
    ComboNote.AddItem "Sans notes"
    For vCount = 0 To 9
        ComboNote.AddItem vCount
    Next
    ComboNote.Text = "1 --> 9"
End Sub

Private Sub TxtActeurs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtGenre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtNum_Change()
    On Error Resume Next
    FrmListe.Liste(0).ListIndex = Int(TxtNum.Text) - 1
    TxtTitre.Text = FrmListe.Liste(1).Text
    TxtGenre.Text = FrmListe.Liste(2).Text
    TxtActeurs.Text = FrmListe.Liste(3).Text
    ComboNote.Text = FrmListe.Liste(4).Text
    TxtPrete.Text = FrmListe.Liste(5).Text
End Sub

Private Sub TxtNum_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        MsgBox "Veuillez n'entrer que des chiffres", vbCritical, "Erreur"
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtPrete_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtTitre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub
