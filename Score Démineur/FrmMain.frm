VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score démineur Loïc - Aldéric"
   ClientHeight    =   8055
   ClientLeft      =   195
   ClientTop       =   465
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Ajouter &victoire"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox LstAlderic 
      Height          =   3180
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.ListBox LstLoic 
      Height          =   3180
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Line GraphAbsAlderic 
      Index           =   0
      X1              =   5160
      X2              =   5520
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line GraphAbsLoic 
      Index           =   0
      X1              =   5160
      X2              =   5520
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line LGraphAlderic 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   5400
      X2              =   5760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line LGraphLoic 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   5400
      X2              =   5760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label LblScoreAbs 
      Height          =   255
      Left            =   6720
      TabIndex        =   40
      Top             =   5040
      Width           =   375
   End
   Begin VB.Shape ShpAbsPos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape ShpPos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "Nombre de mines trouvées:"
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label LblMine 
      Height          =   255
      Left            =   2400
      TabIndex        =   38
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label LblAbsAxeX 
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   37
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label LblAbsAxeY 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   36
      Top             =   6360
      Width           =   255
   End
   Begin VB.Line AbsGradX 
      Index           =   0
      X1              =   4920
      X2              =   4920
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line AbsGradY 
      Index           =   0
      X1              =   4560
      X2              =   4800
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line AbsAxeY 
      X1              =   4680
      X2              =   4680
      Y1              =   3960
      Y2              =   7800
   End
   Begin VB.Line AbsAxeX 
      X1              =   4440
      X2              =   6960
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Lbl3Joueur1 
      Caption         =   "Loïc"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   35
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Lbl3Joueur2 
      Caption         =   "Aldéric"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6960
      TabIndex        =   34
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Evolution du score absolu"
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label LblAxeX 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   32
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label LblAxeY 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   31
      Top             =   3240
      Width           =   255
   End
   Begin VB.Line LGradX 
      Index           =   0
      X1              =   4920
      X2              =   4920
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line LGradY 
      Index           =   0
      X1              =   4560
      X2              =   4800
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label10 
      Caption         =   "Evolution des victoires"
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Parties parfaites (26-25) :"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label LblParfait 
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Lbl2Joueur2 
      Caption         =   "Aldéric"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6720
      TabIndex        =   27
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Lbl2Joueur1 
      Caption         =   "Loïc"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line AxeX 
      X1              =   4440
      X2              =   6960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line AxeY 
      X1              =   4680
      X2              =   4680
      Y1              =   720
      Y2              =   3720
   End
   Begin VB.Label Label9 
      Caption         =   "Score absolu :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Nb victoires - Nb défaites + 5*Nb victoires consécutives + 1/4 de la différence de la plus grosse fritée"
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label LblAbsoluLoic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label LblAbsoluAlderic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Pourcentage :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label LblPourcentLoic 
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label LblPourcentAlderic 
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label LblFriteAlderic 
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label LblFriteLoic 
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Plus grosse fritée :"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label LblNbMatch 
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label LblRecVictConsAlderic 
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label LblRecVictConsLoic 
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label LblVictConsAlderic 
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label LblVictConsLoic 
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Record victoires consécutives :"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Nombre de match :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Victoires consécutives :"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label LblVictAlderic 
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label LblVictLoic 
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Victoires :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label LblJoueur2 
      Caption         =   "Aldéric"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label LblJoueur1 
      Caption         =   "Loïc"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTestLoic As Boolean
Dim vTestAlderic As Boolean

Private Sub CmdAdd_Click()
    FrmAdd.Show
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
Dim vCount As Integer

    FrmFichier.Show
    FrmMain.Hide
End Sub

Private Sub LstLoic_Click()
    If vTestAlderic = False Then
        ShpPos.Top = AxeX.Y1 - 100 * Int(LstLoic.Text) - 50
        ShpPos.Left = AxeY.X1 + LstLoic.ListIndex * 150 + 90
        ShpPos.FillColor = &HFF&
        ShpPos.Visible = True

        ShpAbsPos.Top = AbsAxeX.Y1 - 52 * tTabAbsLoic(LstLoic.ListIndex) - 50
        ShpAbsPos.Left = AbsAxeY.X1 + LstLoic.ListIndex * 150 + 90
        ShpAbsPos.FillColor = &HFF&
        ShpAbsPos.Visible = True
        LblScoreAbs.Caption = Left(Trim(Str(tTabAbsLoic(LstLoic.ListIndex))), 4)
        LblScoreAbs.Left = ShpAbsPos.Left
        LblScoreAbs.ForeColor = &HFF&
        LblScoreAbs.Top = ShpAbsPos.Top - 400
    End If
    vTestLoic = True
    LstAlderic.ListIndex = LstLoic.ListIndex
    vTestAlderic = False
    vTestLoic = False
End Sub

Private Sub LstAlderic_Click()
    If vTestLoic = False Then
        ShpPos.Top = AxeX.Y1 - 100 * Int(LstAlderic.Text) - 50
        ShpPos.Left = AxeY.X1 + LstAlderic.ListIndex * 150 + 90
        ShpPos.FillColor = &HFF0000
        ShpPos.Visible = True

        ShpAbsPos.Top = AbsAxeX.Y1 - 52 * tTabAbsAlderic(LstAlderic.ListIndex) - 50
        ShpAbsPos.Left = AbsAxeY.X1 + LstAlderic.ListIndex * 150 + 90
        ShpAbsPos.FillColor = &HFF0000
        ShpAbsPos.Visible = True
        LblScoreAbs.Caption = Left(Trim(Str(tTabAbsAlderic(LstAlderic.ListIndex))), 4)
        LblScoreAbs.Left = ShpAbsPos.Left
        LblScoreAbs.ForeColor = &HFF0000
        LblScoreAbs.Top = ShpAbsPos.Top - 400
    End If
    vTestAlderic = True
    LstLoic.ListIndex = LstAlderic.ListIndex
    vTestLoic = False
    vTestAlderic = False
End Sub
