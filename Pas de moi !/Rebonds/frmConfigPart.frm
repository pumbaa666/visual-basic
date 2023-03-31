VERSION 5.00
Begin VB.Form frmConfigPart 
   Caption         =   "Configuration de la particule"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtV 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtU 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtY 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtX 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdValider 
      Caption         =   "Valider"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblVecteur 
      Caption         =   "Vecteur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label lblV 
      Caption         =   "V ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label lblU 
      Caption         =   "U ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label lblY 
      Caption         =   "Y ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   420
   End
   Begin VB.Label lblX 
      Caption         =   "X ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lblCoord 
      Caption         =   "Coordonnées :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1770
   End
End
Attribute VB_Name = "frmConfigPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
 txtX.Text = frmSimulation.X
 txtY.Text = frmSimulation.Y
 If frmSimulation.U = 0 And frmSimulation.V = 0 Then
   Do
     frmSimulation.U = Int(20 * Rnd) - 12
   Loop Until frmSimulation.U
   Do
    frmSimulation.V = Int(20 * Rnd) - 12
   Loop Until frmSimulation.V
   txtU.Text = frmSimulation.U
   txtV.Text = frmSimulation.V
 Else
    txtU.Text = frmSimulation.U
    txtV.Text = frmSimulation.V
 End If
End Sub


Private Sub cmdValider_Click()
 txtX.Text = Val(txtX.Text)
 txtY.Text = Val(txtY.Text)
 txtU.Text = Val(txtU.Text)
 txtV.Text = Val(txtV.Text)
 
 frmSimulation.X = Val(txtX.Text)
 frmSimulation.Y = Val(txtY.Text)
 frmSimulation.U = Val(txtU.Text)
 frmSimulation.V = Val(txtV.Text)
 
 Call DessinerTrait(1)
 frmSimulation.Plan.PSet (frmSimulation.X, frmSimulation.Y), vbRed
End Sub

Private Sub cmdFermer_Click()
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
 If frmSimulation.U = 0 And frmSimulation.V = 0 Then
  MsgBox "Vecteur inexistant !!!" & vbCrLf & " Veuillez en choisir un..."
  Cancel = True
 End If
End Sub
