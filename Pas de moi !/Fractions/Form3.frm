VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   7875
   ClientLeft      =   4365
   ClientTop       =   1710
   ClientWidth     =   10680
   LinkTopic       =   "Form3"
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
      Left            =   4560
      TabIndex        =   17
      Top             =   4560
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
      Left            =   4560
      TabIndex        =   16
      Top             =   5040
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
      Left            =   960
      TabIndex        =   15
      Text            =   "Entrez vos fractions  et votre r�ponse. Puis validez. Sinon Quitter."
      Top             =   1080
      Width           =   7815
   End
   Begin VB.CommandButton cmdResultat 
      Caption         =   "Voir le R�sultat ->"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   4680
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
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Text            =   "1"
      Top             =   2520
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
      Left            =   7320
      TabIndex        =   12
      Text            =   "0"
      Top             =   2160
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
      Left            =   5160
      TabIndex        =   10
      Text            =   "1"
      Top             =   2520
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
      Left            =   5160
      TabIndex        =   9
      Text            =   "1"
      Top             =   2160
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
      Left            =   3000
      TabIndex        =   4
      Text            =   "1"
      Top             =   2160
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
      Left            =   8040
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton CmdA 
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
      Left            =   9000
      TabIndex        =   2
      Top             =   2280
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
      Left            =   840
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
      Left            =   3000
      TabIndex        =   0
      Text            =   "1"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Votre r�ponse ici et Valider"
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
      Left            =   6840
      TabIndex        =   20
      Top             =   1800
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
      Left            =   3840
      TabIndex        =   19
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " Apr�s les cases remplies et validation, vous pouvez cliquer sur Voir le R�sultat ."
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
      Left            =   3120
      TabIndex        =   18
      Top             =   3840
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
      Left            =   6600
      TabIndex        =   11
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "+"
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
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "ADDITION DE FRACTIONS"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Num�rateur, nombre entier :"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "D�nominateur,nombre entier :"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdA_Click()
ForeColor = &H0&
Dim nu1, den1, nu2, den2, Sn, Sd, r, q, tr As Double

nu1 = (TxtN1.Text): den1 = (TxtD1.Text)
nu2 = Val(TxtN2.Text): den2 = Val(TxtD2.Text)

'calcul des termes de la somme
Sn = nu1 * den2 + nu2 * den1
Sd = den1 * den2
If TxtD1.Text = 0 Or TxtD2.Text = 0 Or TxtD3.Text = 0 Then
Dim av As String
av = MsgBox("Vous ne pouvez entrer un d�nominateur nul ! Quitter le programme !", vbOK + vbCritical)
If av = vbOK Then End
If av <> vbOK Then End
End If


If Sn = 0 Then TxtRn.Text = 0: txtRd.Text = 1: GoTo dialogue

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
If Sn = 0 Then txtdial = "La fraction est �gale � 0"
If TxtN3 = " " Or TxtD3 = " " Then txtdial.ForeColor = &HFF&: txtdial.Text = "Vous devez entrer un r�sultat num�rique": GoTo dialogue
If TxtN3.Text = TxtRn.Text And TxtD3.Text = txtRd.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "R�sultat exact. Vous pouvez v�rifier."
End If

If TxtN3.Text * txtRd.Text = TxtD3.Text * TxtRn.Text And TxtN3.Text <> TxtRn.Text And TxtD3.Text <> txtRd.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "R�sultat juste mais non irr�ductible. Continuez ou voir R�sultat."
End If
If TxtN3.Text * txtRd.Text <> TxtD3.Text * TxtRn.Text Then
txtdial.Text = " "
txtdial.ForeColor = &HFF&
txtdial.Text = "R�sultat faux. Continuez ou voir R�sultat."
End If

fin:



End Sub

Private Sub Cmdquitter3_Click()
End
End Sub

Private Sub cmdResultat_Click()

TxtRn.Visible = True
txtRd.Visible = True
txtdial.ForeColor = &H0&
txtdial.Text = "Entrez votre fraction et votre r�ponse. Puis validez. Sinon Quitter."
End Sub

Private Sub CmdRetour_Click()
Form6.Hide
Form5.Hide
Form4.Hide
Form3.Hide
Form2.Hide
Form1.Show
End Sub

