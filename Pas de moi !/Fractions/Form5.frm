VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   7875
   ClientLeft      =   5190
   ClientTop       =   1980
   ClientWidth     =   10830
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10830
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
      Height          =   360
      Left            =   2760
      TabIndex        =   12
      Text            =   "1"
      Top             =   2400
      Width           =   1695
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
      Left            =   720
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton CmdM 
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
      Left            =   8880
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Cmdquitter3 
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
      Height          =   615
      Left            =   8040
      TabIndex        =   9
      Top             =   6480
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
      Left            =   2760
      TabIndex        =   8
      Text            =   "1"
      Top             =   2040
      Width           =   1695
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
      Height          =   360
      Left            =   4800
      TabIndex        =   7
      Text            =   "1"
      Top             =   2040
      Width           =   1455
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
      Left            =   4800
      TabIndex        =   6
      Text            =   "1"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox TxtN3 
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
      Left            =   6960
      TabIndex        =   5
      Text            =   "0"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox TxtD3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   4
      Text            =   "1"
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton cmdResultat 
      Caption         =   "Voir le Résultat ->"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
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
      Left            =   600
      TabIndex        =   2
      Text            =   "Entrez vos fractions  et votre réponse. Puis validez. Sinon Quitter."
      Top             =   960
      Width           =   7815
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
      Left            =   4200
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
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
      Height          =   435
      Left            =   4200
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
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
      Left            =   6480
      TabIndex        =   20
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Dénominateur,nombre entier :"
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
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Numérateur, nombre entier :"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "MULTIPLICATION  DE FRACTIONS"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   2160
      Width           =   615
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
      Left            =   2760
      TabIndex        =   14
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "< Entrez vos 2 fractions >"
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
      Left            =   3480
      TabIndex        =   13
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdM_Click()
ForeColor = &H0&
Dim nu1, den1, nu2, den2, Sn, Sd, r, q, tr As Double

nu1 = (TxtN1.Text): den1 = (TxtD1.Text)
nu2 = Val(TxtN2.Text): den2 = Val(TxtD2.Text)

'calcul des termes de la multipliction
Sn = nu1 * nu2
Sd = den1 * den2

If TxtD1.Text = 0 Or TxtD2.Text = 0 Or TxtD3.Text = 0 Then
Dim av As String
av = MsgBox("Vous ne pouvez entrer un dénominateur nul ! Quitter le programme !", vbOK + vbCritical)
If av = vbOK Then End
If av <> vbOK Then End
End If

If Sn = 0 Then TxtRn = 0: txtRd = 1: GoTo dialogue

If Sn = 0 Then TxtRn = 0: txtRd = 1: GoTo dialogue
'simplification
sna = Abs(Sn): sda = Abs(Sd): H = sna: G = sda

If sna < sda Then tr = sna: sna = sda: sda = tr
divise:
r = sna - sda * Int(sna / sda)
If r = 0 Then d = sda: GoTo pgcd
If r <> 0 Then sna = sda: sda = r: GoTo divise
pgcd:
TxtRn.Visible = False
txtRd.Visible = False
If (Sn / Sd) < 0 Then H = -H
TxtRn.Text = H / d: txtRd.Text = G / d


dialogue:
If Sn = 0 Then txtdial.Text = "La fraction vaut 0"
If TxtN3.Text = TxtRn.Text And TxtD3.Text = txtRd.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "Résultat exact. Vous pouvez vérifier."
End If

If TxtN3.Text * txtRd.Text = TxtD3.Text * TxtRn.Text And TxtN3.Text <> TxtRn.Text And TxtD3.Text <> txtRd.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "Résultat juste mais non irréductible. Continuez ou voir Résultat."
End If
If TxtN3.Text * txtRd.Text <> TxtD3.Text * TxtRn.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "Résultat faux. Continuez ou voir Résultat."
End If





End Sub

Private Sub Cmdquitter3_Click()
End
End Sub

Private Sub cmdResultat_Click()
ForeColor = &H0&
TxtRn.Visible = True
txtRd.Visible = True
txtdial.ForeColor = &H0&
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
