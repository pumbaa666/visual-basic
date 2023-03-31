VERSION 5.00
Begin VB.Form frmGestionTrait 
   Caption         =   "Gestion des traits"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer ce trait"
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter ce trait"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdValider 
      Caption         =   "Valider"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtAB 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   15
      Text            =   "0"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtAB 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Text            =   "0"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtCoord 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   13
      Text            =   "0"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtCoord 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   12
      Text            =   "0"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtCoord 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Text            =   "0"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtCoord 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      Text            =   "0"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtSelection 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox opt4Bord 
      Caption         =   "4 Bords = Traits ?"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Option :"
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
      Left            =   360
      TabIndex        =   18
      Top             =   840
      Width           =   960
   End
   Begin VB.Label lblB 
      Caption         =   "b ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   9
      Top             =   3240
      Width           =   330
   End
   Begin VB.Label lblA 
      Caption         =   "a ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   330
   End
   Begin VB.Label lblY2 
      Caption         =   "Y2 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   7
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label lblY1 
      Caption         =   "Y1 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label lblN 
      Caption         =   "N ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label lblM 
      Caption         =   "M ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblSelection 
      Caption         =   "Sélection du trait n°           / 4"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   3810
   End
   Begin VB.Label lblTitreGestion 
      Alignment       =   2  'Center
      Caption         =   "Gestion des traits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4965
   End
End
Attribute VB_Name = "frmGestionTrait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim LimAppel As Byte








'Affichage Données

Private Sub Form_Load()
  If LimAppel = 0 Then _
  txtSelection.Text = frmSimulation.nbrTrait + 4 * (frmSimulation.OptBord)
 LimAppel = 1
  txtCoord(0).Text = lstT(Val(txtSelection.Text)).M
  txtCoord(1).Text = lstT(Val(txtSelection.Text)).Y1
  txtCoord(2).Text = lstT(Val(txtSelection.Text)).N
  txtCoord(3).Text = lstT(Val(txtSelection.Text)).Y2
  txtAB(0).Text = lstT(Val(txtSelection.Text)).a
  txtAB(1).Text = lstT(Val(txtSelection.Text)).b
  lblSelection.Caption = "Sélection du trait n°           / " & frmSimulation.nbrTrait + 4 * (OptBord)
 LimAppel = 0
End Sub

'OPTION 4 Bords


Private Sub opt4Bord_Click()
 frmSimulation.OptBord = opt4Bord.Value
 frmSimulation.nbrTrait = frmSimulation.nbrTrait - 4 - 8 * frmSimulation.OptBord
 If opt4Bord.Value Then Call CreerBord(frmSimulation.nbrTrait)
 lblSelection.Caption = "Sélection du trait n°           / " & frmSimulation.nbrTrait + 4 * (OptBord)
 If frmSimulation.nbrTrait Then cmdValider.Enabled = True Else cmdValider.Enabled = False
 ReDim Preserve lstT(frmSimulation.nbrTrait)
 Call txtSelection_Change
 Call DessinerTrait(1)
End Sub













'Selection

Private Sub txtSelection_Click()
 txtSelection.SelStart = 0
 txtSelection.SelLength = Len(txtSelection)
End Sub
Private Sub txtSelection_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtSelection_Change()
If LimAppel = 0 Then
 If Val(txtSelection.Text) > frmSimulation.nbrTrait Then txtSelection.Text = frmSimulation.nbrTrait
 LimAppel = 1
 Call Form_Load
End If
End Sub



'AJOUTER / SUPPRIMER


Private Sub cmdAjouter_Click()

 If Val(txtCoord(0).Text) > Val(txtCoord(2).Text) Then
    MsgBox "Veuillez choisir une valeur de M plus petite ou égale à N"
    Exit Sub
 End If

 frmSimulation.nbrTrait = frmSimulation.nbrTrait + 1
 ReDim Preserve lstT(frmSimulation.nbrTrait)
 
 Dim z As Integer
 If frmSimulation.OptBord = True Then
    For z = Val(frmSimulation.nbrTrait) To Val(frmSimulation.nbrTrait) - 3 Step -1
        lstT(z).M = lstT(z - 1).M
        lstT(z).N = lstT(z - 1).N
        lstT(z).Y1 = lstT(z - 1).Y1
        lstT(z).Y2 = lstT(z - 1).Y2
        lstT(z).a = lstT(z - 1).a
        lstT(z).b = lstT(z - 1).b
    Next z
 End If
 
 lstT(z).M = Val(txtCoord(0).Text)
 lstT(z).N = Val(txtCoord(2).Text)
 lstT(z).Y1 = Val(txtCoord(1).Text)
 lstT(z).Y2 = Val(txtCoord(3).Text)
 lstT(z).a = Val(Replace(txtAB(0).Text, ",", "."))
 lstT(z).b = Val(Replace(txtAB(1).Text, ",", "."))
 txtSelection.Text = frmSimulation.nbrTrait - 4 * (opt4Bord)
 lblSelection.Caption = "Sélection du trait n°           / " & frmSimulation.nbrTrait + 4 * (OptBord)
 If frmSimulation.nbrTrait Then
    cmdValider.Enabled = True
    cmdSupprimer.Enabled = True
 Else
    cmdValider.Enabled = False
    cmdSupprimer.Enabled = False
 End If
 Call DessinerTrait(1)
End Sub

Private Sub cmdValider_Click()
 If Val(txtCoord(0).Text) < Val(txtCoord(2).Text) And LimAppel = 0 Then
    lstT(Val(txtSelection.Text)).M = Val(txtCoord(0).Text)
    lstT(Val(txtSelection.Text)).N = Val(txtCoord(2).Text)
    lstT(Val(txtSelection.Text)).Y1 = Val(txtCoord(1).Text)
    lstT(Val(txtSelection.Text)).Y2 = Val(txtCoord(3).Text)
    lstT(Val(txtSelection.Text)).a = Val(Replace(txtAB(0).Text, ",", "."))
    lstT(Val(txtSelection.Text)).b = Val(Replace(txtAB(1).Text, ",", "."))
    Call DessinerTrait(1)
 Else
    MsgBox "Veuillez choisir une valeur de M plus petite ou égale à N"
 End If
End Sub

Private Sub cmdSupprimer_Click()
 
 frmSimulation.nbrTrait = frmSimulation.nbrTrait - 1
 Dim z As Byte
 For z = Val(txtSelection.Text) To frmSimulation.nbrTrait
     lstT(z).M = lstT(z + 1).M
     lstT(z).N = lstT(z + 1).N
     lstT(z).Y1 = lstT(z + 1).Y1
     lstT(z).Y2 = lstT(z + 1).Y2
     lstT(z).a = lstT(z + 1).a
     lstT(z).b = lstT(z + 1).b
 Next z
 ReDim Preserve lstT(frmSimulation.nbrTrait)

 lblSelection.Caption = "Sélection du trait n°           / " & frmSimulation.nbrTrait + 4 * (OptBord)
 If frmSimulation.nbrTrait Then
    cmdValider.Enabled = True
    cmdSupprimer.Enabled = True
 Else
    cmdValider.Enabled = False
    cmdSupprimer.Enabled = False
 End If
 Call txtSelection_Change
 Call DessinerTrait(1)
End Sub





Private Sub txtCoord_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtCoord_Change(Index As Integer)
 If Val(txtCoord(0).Text) < Val(txtCoord(2).Text) And LimAppel = 0 Then
    LimAppel = 1
    txtAB(0).Text = (Val(txtCoord(3).Text) - Val(txtCoord(1).Text)) / (Val(txtCoord(2).Text) - Val(txtCoord(0).Text))
    txtAB(1).Text = Val(txtCoord(1).Text) - Val(Replace(txtAB(0).Text, ",", ".")) * Val(txtCoord(0).Text)
    LimAppel = 0
 End If
End Sub


Private Sub txtAB_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii > 57 Or KeyAscii < 48 And KeyAscii <> 8 Then
  Select Case Chr(KeyAscii)
       Case ".", ","   ' Si vous voulez que "," soit consideré comme "."
            KeyAscii = Asc(".")
            If (InStr(1, txtAB(Index), ".") Or Mid(txtAB(Index), txtAB(Index).SelStart + 1, 1) = "-") _
               And Not (InStr(1, txtAB(Index).SelText, ".")) Then KeyAscii = 0
       Case "-"
       Dim MEM As Byte
       MEM = txtAB(Index).SelStart
       If InStr(1, txtAB(Index), "-") Then
           If Not (InStr(1, txtAB(Index).SelText, "-")) Then
       'Cette ligne  met SelStart à 0 donc vala pk on l'a memorisé avant !
           txtAB(Index) = Mid(txtAB(Index), 2, txtAB(Index).SelStart - 1) _
               & Right(txtAB(Index), Len(txtAB(Index)) - txtAB(Index).SelStart)
           txtAB(Index).SelStart = MEM - 1
           KeyAscii = 0
           End If
       Else
           txtAB(Index) = "-" & txtAB(Index)
           txtAB(Index).SelStart = MEM + 1
           KeyAscii = 0
       End If
  Case Else
       KeyAscii = 0
       End Select
 End If
End Sub
Private Sub txtAB_Change(Index As Integer)
 If LimAppel = 0 Then
    LimAppel = 1
    If Index Then
       txtCoord(1).Text = Val(Replace(txtAB(0).Text, ",", ".")) * txtCoord(0).Text + Val(Replace(txtAB(1).Text, ",", "."))
       txtCoord(3).Text = Val(Replace(txtAB(0).Text, ",", ".")) * txtCoord(2).Text + Val(Replace(txtAB(1).Text, ",", "."))
    Else
       txtCoord(3).Text = Val(Replace(txtAB(0).Text, ",", ".")) * Val(txtCoord(2).Text) + Val(Replace(txtAB(1).Text, ",", "."))
    End If
    LimAppel = 0
 End If
End Sub



Private Sub cmdFermer_Click()
 Unload Me
End Sub
