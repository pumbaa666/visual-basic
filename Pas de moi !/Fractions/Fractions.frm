VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   4365
   ClientTop       =   1710
   ClientWidth     =   10725
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10725
   Begin VB.OptionButton OptSou 
      BackColor       =   &H8000000D&
      Caption         =   "Soustraction de fractions"
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
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton CmdFin 
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
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Cmdmenu 
      Caption         =   "OK"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H8000000D&
      Caption         =   "Division de fractions"
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
      Left            =   720
      TabIndex        =   3
      Top             =   4080
      Width           =   2775
   End
   Begin VB.OptionButton optM 
      BackColor       =   &H8000000D&
      Caption         =   "Multiplication de fractions"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
   End
   Begin VB.OptionButton optS 
      BackColor       =   &H8000000D&
      Caption         =   "Simplifier une fraction"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H8000000D&
      Caption         =   "Addition de  fractions"
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
      Height          =   555
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "A L' ATTENTION DES ELEVES :"
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
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Entrez des entiers relatifs. Ex. : -1,2/3 s'écrit : -12/30 . On peut enchaîner un résultat  avec trois ou quatre fractions, etc..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   5640
      Width           =   8775
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "Vous devrez obligatoirement entrer une réponse (qui sera appréciée), avant  de pouvoir lire le résultat."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   5400
      Width           =   8895
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Sanction : votre travail sera annulé, vous devrez sortir du programme et tout recommencer  ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   5040
      Width           =   8415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Enfin on rappelle que : si un des termes est entier son dénominateur est 1 ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   6240
      Width           =   8415
   End
   Begin VB.Label lblAv 
      BackColor       =   &H80000005&
      Caption         =   "Vous ne devez pas entrer de dénominateur nul ! Ex.: 8/0 = q donne q x 0 = 8, ce qui est impossible . "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   4800
      Width           =   9015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cochez votre choix et validez par OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   855
      Left            =   5520
      TabIndex        =   8
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPERATIONS SUR LES FRACTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub





Private Sub Cmdmenu_Click()


If optS.Value = True Then
Form1.Hide
Form2.Show
Form3.Hide
Form4.Hide
Form6.Hide
Form5.Hide
End If

If optA.Value = True Then
Form1.Hide
Form2.Hide
Form3.Show
Form4.Hide
Form6.Hide
Form5.Hide
End If

If OptSou.Value = True Then
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Show
Form6.Hide
Form5.Hide
End If

If optM.Value = True Then
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form6.Hide
Form5.Show
End If

If optD.Value = True Then
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Show
End If
End Sub

Private Sub CmdFin_Click()
End
End Sub

