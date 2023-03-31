VERSION 5.00
Begin VB.Form FrmPreter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mes DVD - Prêter un DVD"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   4455
      Begin VB.TextBox TxtNom 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Prêter un DVD à :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   3975
      Begin VB.Timer ClkWait 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox TxtNum 
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton CmdAnnuler 
         Caption         =   "&Annuler"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Numéro : "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Shape ShpOk 
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3480
         Shape           =   3  'Circle
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Timer ClkOk 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   720
   End
End
Attribute VB_Name = "FrmPreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClkOk_Timer()
Static vCntOk As Boolean
    If vCntOk = False Then
        vCntOk = True
        ShpOk.FillColor = &HFF00&
        ClkOk.Interval = 500
    Else
        ShpOk.FillColor = &HFF&
        ClkOk.Enabled = False
        ClkOk.Interval = 1
        vCntOk = False
    End If
End Sub

Private Sub ClkWait_Timer()
Static vCntWait As Boolean
    If vCntWait = False Then
        FrmWait.Show
        ClkOk.Enabled = True
        vCntWait = True
    Else
        DelDVD TxtNum.Text
        TxtNum.Text = ""
        TxtNum.SetFocus
        FrmWait.Hide
        FrmListe.Liste(0).ListIndex = 0
        ClkWait.Enabled = False
        vCntWait = False
    End If
End Sub

Private Sub CmdAnnuler_Click()
    FrmPreter.Hide
End Sub

Private Sub CmdOk_Click()
Dim vCount As Integer
    If Mid(FrmPreter.Caption, 11, 5) = "Prête" Then
        If TxtNom.Text = "" Then
            TxtNom.SetFocus
            MsgBox "Veuillez indiquer à qui vous le prêtez", vbCritical, "Erreur"
        ElseIf TxtNum.Text = "" Then
            TxtNum.SetFocus
            MsgBox "Il manque le numéro du DVD", vbCritical, "Erreur"
        ElseIf Int(TxtNum.Text) > vNbDVDTot Or Int(TxtNum.Text) = 0 Then
            MsgBox "Ce DVD n'existe pas", vbCritical, "Erreur"
            TxtNum.Text = ""
            TxtNum.SetFocus
        Else
            FrmListe.Liste(5).ListIndex = Int(TxtNum.Text) - 1
            If FrmListe.Liste(5).Text <> "" Then
                MsgBox "Ce DVD est déjà prêté", vbCritical, "Erreur"
            Else
                tListe(5, Int(TxtNum.Text) - 1) = TxtNom.Text
                Preter Int(TxtNum.Text), TxtNom.Text
                RefreshListePrete
                ClkOk.Enabled = True
            End If
            TxtNum.Text = ""
            TxtNum.SetFocus
        End If
    ElseIf Mid(FrmPreter.Caption, 11, 5) = "Repre" Then
        If TxtNum.Text = "" Then
            TxtNum.SetFocus
            MsgBox "Il manque le numéro du DVD", vbCritical, "Erreur"
        ElseIf Int(TxtNum.Text) > vNbDVDTot Or Int(TxtNum.Text) = 0 Then
            MsgBox "Ce DVD n'existe pas", vbCritical, "Erreur"
            TxtNum.Text = ""
            TxtNum.SetFocus
        Else
            FrmListe.Liste(5).ListIndex = Int(TxtNum.Text) - 1
            If FrmListe.Liste(5).Text = "" Then
                MsgBox "Ce DVD n'était pas prêté", vbCritical, "Erreur"
            Else
                tListe(5, Int(TxtNum.Text)) = ""
                Preter Int(TxtNum.Text), ""
                FrmListe.Liste(5).RemoveItem Int(TxtNum.Text) - 1
                FrmListe.Liste(5).AddItem "", Int(TxtNum.Text) - 1
                FrmListe.Liste(5).ListIndex = Int(TxtNum.Text) - 1
                ClkOk.Enabled = True
            End If
            TxtNum.Text = ""
            TxtNum.SetFocus
        End If
    ElseIf Mid(FrmPreter.Caption, 11, 5) = "Attei" Then
        If TxtNum.Text = "" Then
            TxtNum.SetFocus
            MsgBox "Il manque le numéro du DVD", vbCritical, "Erreur"
        ElseIf Int(TxtNum.Text) > vNbDVDTot Or Int(TxtNum.Text) = 0 Then
            MsgBox "Ce DVD n'existe pas", vbCritical, "Erreur"
            TxtNum.Text = ""
            TxtNum.SetFocus
        Else
            FrmListe.Liste(5).ListIndex = vNbDVDTot - 1
            FrmListe.Liste(5).ListIndex = Int(TxtNum.Text) - 1
            TxtNum.Text = ""
            TxtNum.SetFocus
        End If
    ElseIf Mid(FrmPreter.Caption, 11, 5) = "Suppr" Then
        If TxtNum.Text = "" Then
            TxtNum.SetFocus
            MsgBox "Il manque le numéro du DVD", vbCritical, "Erreur"
        ElseIf Int(TxtNum.Text) > vNbDVDTot Or Int(TxtNum.Text) = 0 Then
            MsgBox "Ce DVD n'existe pas", vbCritical, "Erreur"
            TxtNum.Text = ""
            TxtNum.SetFocus
        Else
            ClkWait.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    TxtNom.Text = ""
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
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
