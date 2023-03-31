VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajouter Entrée"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkInc 
      Caption         =   "Incrémenter le nom"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox TxtDuree 
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox TxtArrivee 
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox TxtNom 
      Height          =   285
      Left            =   960
      MaxLength       =   5
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Durée"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Arrivée"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nom"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
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
Dim vCount As Integer
Dim vExiste As Integer
Dim vNum As Integer

    If TxtNom.Text <> "" And TxtArrivee.Text <> "" And TxtDuree.Text <> "" Then
        
        If FrmMain.LstNom.ListCount <> 0 Then
            For vCount = 0 To FrmMain.LstNom.ListCount - 1
                FrmMain.LstNom.ListIndex = vCount
                FrmMain.LstArrivee.ListIndex = vCount
                If FrmMain.LstNom.Text = TxtNom.Text Or FrmMain.LstArrivee.Text = TxtArrivee.Text Then
                    MsgBox "Donnée redondante", vbCritical, "Erreur"
                    vExiste = True
                    Exit For
                End If
            Next
        End If
        
        If vExiste = False Then
            FrmMain.LstNom.AddItem TxtNom.Text
            FrmMain.LstArrivee.AddItem TxtArrivee.Text
            FrmMain.LstDuree.AddItem TxtDuree.Text

            ReDim Preserve tEntree(vNbEntree)
            tEntree(vNbEntree).vNom = TxtNom.Text
            tEntree(vNbEntree).vArrivee = Int(TxtArrivee.Text)
            tEntree(vNbEntree).vDuree = Int(TxtDuree.Text)
            tEntree(vNbEntree).vMin = False

            vNbEntree = vNbEntree + 1
            If ChkInc.Value = Checked Then
                If IsNumeric(Right(TxtNom.Text, 1)) Then
                    vNum = Right(TxtNom.Text, 1)
                    vNum = vNum + 1
                    If vNum = 10 Then
                        vNum = 0
                    End If
                    TxtNom.Text = Left(TxtNom.Text, Len(TxtNom.Text) - 1) & vNum
                Else
                    TxtNom.Text = Left(TxtNom.Text, Len(TxtNom.Text)) & "1"
                End If
            End If
        End If
    Else
        MsgBox "Il manque des paramètres", vbCritical, "Erreur"
    End If
End Sub

Function TestKeyPress(vCar As Integer) As Boolean
    If (vCar < 48 Or vCar > 57) And vCar <> 8 And vCar <> 13 Then
        TestKeyPress = False
    Else
        TestKeyPress = True
    End If

    If vCar = 13 Then
        CmdOk_Click
        TestKeyPress = True
    End If
End Function

Private Sub TxtArrivee_KeyPress(KeyAscii As Integer)
    If TestKeyPress(KeyAscii) = False Then
        MsgBox "N'entrez que des chiffres", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDuree_KeyPress(KeyAscii As Integer)
    If TestKeyPress(KeyAscii) = False Then
        MsgBox "N'entrez que des chiffres", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        MsgBox "Caractère interdit", vbCritical, "Erreur"
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub
