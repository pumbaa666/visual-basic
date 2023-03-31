VERSION 5.00
Begin VB.Form FrmVit 
   Caption         =   "Vitesse de déplacement"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtVit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "500"
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Choisissez la vitesse du moteur (100 -> 999)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "FrmVit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vVit As Integer
Dim vNbVirVit As Boolean

Private Sub CmdAnnuler_Click()
    TxtVit.Text = vVit
    FrmMain.Show
    FrmVit.Hide
End Sub

Private Sub CmdOk_Click()
    If TxtVit.Text = "" Then
        MsgBox "Veuillez entrer une valeur!!!", vbCritical, "Erreur"
    ElseIf Int(TxtVit.Text) < 100 Or Int(TxtVit.Text) > 999 Then
        MsgBox "La vitesse doit être comprise entre 0 et 999!!!", vbCritical, "Erreur"
    Else
        
' Change la vitesse du moteur
        vChaine = "vit " & TxtVit.Text & ";"
        FrmMain.ClkEnvoie.Enabled = True
        FrmMain.Show
        FrmVit.Hide
    End If
End Sub

Private Sub Form_Load()
' Sauvegarde la vitesse dans une variable pour la réstaurer si on click sur Annuler
    vVit = Int(TxtVit.Text)
End Sub

Private Sub TxtVit_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 13) Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur de saisie"
        KeyAscii = 0
    End If
     
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
    
End Sub
