VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   7875
   ClientLeft      =   5055
   ClientTop       =   2250
   ClientWidth     =   10680
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10680
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
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
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
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
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
      TabIndex        =   10
      Text            =   "Entrez vos fractions  et votre réponse. Puis validez. Sinon Quitter."
      Top             =   840
      Width           =   7815
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
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
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
      TabIndex        =   8
      Text            =   "1"
      Top             =   2310
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
      TabIndex        =   7
      Text            =   "0"
      Top             =   1920
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
      Top             =   2280
      Width           =   1455
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
      TabIndex        =   5
      Text            =   "1"
      Top             =   1920
      Width           =   1455
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
      TabIndex        =   4
      Text            =   "1"
      Top             =   1920
      Width           =   1695
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
      Height          =   735
      Left            =   7560
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton CmdD 
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
      Left            =   8640
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
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
      Left            =   1560
      TabIndex        =   1
      Top             =   6480
      Width           =   1815
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
      Height          =   360
      Left            =   2760
      TabIndex        =   0
      Text            =   "1"
      Top             =   2280
      Width           =   1695
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
      Left            =   6600
      TabIndex        =   20
      Top             =   1560
      Width           =   2415
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
      TabIndex        =   19
      Top             =   1560
      Width           =   2295
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
      Left            =   2880
      TabIndex        =   18
      Top             =   3480
      Width           =   4455
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
      Left            =   6240
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "DIVISION  DE FRACTIONS"
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
      Left            =   1800
      TabIndex        =   15
      Top             =   0
      Width           =   5655
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
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
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
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdD_Click()
ForeColor = &H0&
Dim nu1, den1, nu2, den2, Sn, Sd, r, q, tr As Double

nu1 = (TxtN1.Text): den1 = (TxtD1.Text)
nu2 = Val(TxtN2.Text): den2 = Val(TxtD2.Text)

'calcul des termes de la division
Sn = nu1 * den2
Sd = den1 * nu2
If TxtD1.Text = 0 Or TxtD2.Text = 0 Or TxtD3.Text = 0 Or TxtN2.Text = 0 Then
Dim av As String
av = MsgBox("Ici : a/b :c/d = a/b x d/c. Vous ne pouvez entrer un dénominateur nul ! Quitter le programme !", vbOK + vbCritical)
If av = vbOK Then End
If av <> vbOK Then End
End If

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

