VERSION 5.00
Begin VB.Form FrmOption 
   Caption         =   "Nombre de pas"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4170
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton CmdValider 
      Caption         =   "&Valider"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TxtPas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "48"
      Top             =   240
      Width           =   615
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblDeg 
      Caption         =   "Angle entre chaque pas : 3.75°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de pas du moteur:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNbPas As Integer

Private Sub CmdAnnuler_Click()
    TxtPas.Text = vVit
    FrmMain.Show
    FrmOption.Hide
End Sub

Private Sub CmdOk_Click()
    If TxtPas.Text = "" Then
        MsgBox "Entrez une valeur!", vbCritical, "Erreur"
    ElseIf TxtPas.Text = "0" Then
        MsgBox "Entrez une valeur différente de zéro!", vbCritical, "Erreur"
    Else
        vPas = (Round(360 / Int(TxtPas.Text), 1)) / 2
        LblDeg.Caption = "Angle entre chaque pas :" & Str(vPas) & "°"
        
' Initialise la position du moteur si on change le nb de pas
        vChaine = "ini;"
        FrmMain.ClkEnvoie.Enabled = True
        FrmMain.Show
        FrmOption.Hide
    End If
End Sub

Private Sub CmdValider_Click()
    If TxtPas.Text = "" Then
        MsgBox "Entrez une valeur!", vbCritical, "Erreur"
    ElseIf TxtPas.Text = "0" Then
        MsgBox "Entrez une valeur différente de zéro!", vbCritical, "Erreur"
    Else
        vPas = (Round(360 / Int(TxtPas.Text), 1)) / 2
        LblDeg.Caption = "Angle entre chaque pas :" & Str(vPas) & "°"

' Initialise la position du moteur si on change le nb de pas
        vChaine = "ini;"
        FrmMain.ClkEnvoie.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    vVit = TxtPas.Text
End Sub

Private Sub Txtpas_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 13 Then
            MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur de saisie"
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        CmdValider_Click
    End If
End Sub
