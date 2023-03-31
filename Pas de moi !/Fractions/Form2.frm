VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   7875
   ClientLeft      =   4365
   ClientTop       =   1710
   ClientWidth     =   10710
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10710
   Begin VB.CommandButton CmdResultat 
      Caption         =   "Voir le résultat -> "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtRd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox TxtRn 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   12
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtdial 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      TabIndex        =   11
      Text            =   "Entrez votre fraction et votre réponse. Puis validez. Sinon Quitter."
      Top             =   1320
      Width           =   7815
   End
   Begin VB.TextBox TxtD2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Text            =   "1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox TxtN2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6360
      TabIndex        =   8
      Text            =   "0"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox TxtD1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Text            =   "1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton CmdRetour 
      Caption         =   "Retour au Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton CmdS 
      Caption         =   "Valider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Cmdquitter2 
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox TxtN1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Text            =   "1"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Votre réponse ici et Valider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "<Votre fraction >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " Après les cases remplies et validation, vous pouvez cliquer sur Voir le Résultat ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3480
      TabIndex        =   15
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Dénominateur, nombre entier :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Numérateur, nombre  entier :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "SIMPLIFICATION D'UNE FRACTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Cmdquitter2_Click()
End
End Sub

Private Sub cmdResultat_Click()
txtdial.ForeColor = &H0&
TxtRn.Visible = True
txtRd.Visible = True
txtdial.Text = "Entrez votre fraction et votre réponse. Puis validez. Sinon Quitter."
End Sub

Private Sub CmdRetour_Click()
Form6.Hide
Form5.Hide
Form4.Hide
Form3.Hide
Form2.Hide
Form1.Show
End Sub

Private Sub CmdS_Click()
debut:
 Dim nu, den, r, q, tr As Double
 ForeColor = &H0&
 
 nu = (TxtN1.Text): den = (TxtD1.Text)
 nua = Abs(TxtN1.Text): dena = Abs(TxtD1.Text): H = nua: G = dena

If TxtD1.Text = 0 Or TxtD2.Text = 0 Then
Dim av As String
av = MsgBox("Vous ne pouvez entrer un dénominateur nul ! Annuler l'exercice ?", vbOK + vbCritical)
If av = vbOK Then End
If av <> vbOK Then End
End If

If nu = 0 Then TxtRn = 0: txtRd = dena: GoTo dialogue

If nua < dena Then tr = nua: nua = dena: dena = tr
250
r = nua - dena * Int(nua / dena)
If r = 0 Then d = dena: GoTo pgcd
If r <> 0 Then nua = dena: dena = r: GoTo 250
pgcd:
TxtRn.Visible = False
txtRd.Visible = False
If (nu / den) < 0 Then H = -H
TxtRn.Text = H / d: txtRd.Text = G / d

dialogue:
If nu = 0 Then txtdial.Text = "La fraction est égale  à 0"

If TxtN2.Text = TxtRn.Text And TxtD2.Text = txtRd.Text Then

txtdial.ForeColor = &HFF&
txtdial.Text = "Résultat exact. Vous pouvez vérifier."
End If

If TxtN2.Text * txtRd.Text = TxtD2.Text * TxtRn.Text And TxtN2.Text <> TxtRn.Text And TxtD2.Text <> txtRd.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "Résultat juste mais non irréductible. Continuez ou voir Résultat."
End If

If TxtN2.Text * txtRd.Text <> TxtD2.Text * TxtRn.Text Then
txtdial.ForeColor = &HFF&
txtdial.Text = "Résultat faux. Continuez ou voir Résultat."
End If

End Sub

Private Sub TextN1_Change()
Dim nu, den, r, q, tr  As Double
 
End Sub



Private Sub Command1_Click()

End Sub

Private Sub TxtD2_Change()
Dim nu, den, r, q, tr As Double
End Sub

Private Sub TxtN2_Change()
Dim nu, den, r, q, tr As Double
End Sub
