VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulation de force de frottement"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerObj 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   4560
   End
   Begin VB.TextBox TxtKg 
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
      Left            =   4440
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdZero 
      Caption         =   "&Remise à zéro"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox Combo 
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
      ItemData        =   "FROTTE~1.frx":0000
      Left            =   840
      List            =   "FROTTE~1.frx":0016
      TabIndex        =   5
      Text            =   "Choisissez le type de matériaux"
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Timer TimerTitre 
      Interval        =   250
      Left            =   240
      Top             =   240
   End
   Begin VB.Line LineTourne 
      BorderWidth     =   2
      X1              =   6120
      X2              =   6360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   6120
      X2              =   6120
      Y1              =   4800
      Y2              =   5040
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6000
      Shape           =   2  'Oval
      Top             =   4680
      Width           =   255
   End
   Begin VB.Line LineTire 
      X1              =   2040
      X2              =   6120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Shape Shape 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   840
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   840
      X2              =   6600
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label LblForce2 
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label LblForce1 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Masse de l'objet en kg"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label LblTitre 
      Caption         =   "Simulation de force de frottement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tMuGen(2) As Single
Dim vForce As Integer

Private Sub CmdZero_Click()
    Shape.Left = 840
    TimerObj.Enabled = False
End Sub

Private Sub Combo_Click()
Dim tMu(6, 2) As Single
    tMu(0, 0) = 0.5
    tMu(1, 0) = 0.6
    tMu(2, 0) = 0.2
    tMu(3, 0) = 0.03
    tMu(4, 0) = 0.1
    tMu(5, 0) = 0.05
    tMu(0, 1) = 0.4
    tMu(1, 1) = 0.4
    tMu(2, 1) = 0.2
    tMu(3, 1) = 0.03
    tMu(4, 1) = 0.03
    tMu(5, 1) = 0.05
    tMuGen(0) = tMu(Combo.ListIndex, 0)
    tMuGen(1) = tMu(Combo.ListIndex, 1)
End Sub

Private Sub Combo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    MsgBox "N'entrez rien dans la combobox!!!", vbCritical, "Simulation de force de frottement"
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdStart_Click()
    If TxtKg.Text = "" Then
        MsgBox "Veuillez entrer une masse", vbCritical, "Simulation de force de frottement"
    ElseIf Int(TxtKg.Text) < 1 Or Int(TxtKg.Text) > 10000 Then
        MsgBox "La masse doit être comprise entre 1 et 10'000 kg!", vbCritical, "Simulation de force de frottement"
        TxtKg.Text = ""
    ElseIf tMuGen(0) = 0 Then
        MsgBox "Veuillez choisir un materiaux", vbCritical, "Simulation de force de frottement"
    Else
        vForce = Int(Int(TxtKg.Text) * 9.81 * tMuGen(1))
        LblForce1.Caption = "La force nécesaire à faire bouger l'objet est de " & Str(Int(Int(TxtKg.Text) * 9.81 * tMuGen(0))) & " N"
        LblForce2.Caption = "La force nécesaire à maintenir l'objet en mouvement est de " & Str(vForce) & " N"
        TimerObj.Enabled = True
    End If
End Sub

Private Sub TimerObj_Timer()
Static vCount As Double
    If Shape.Left < 4750 Then
        If vForce > 1000 Then
            vForce = 950
        End If
        Shape.Left = Shape.Left + (20 - vForce / 50)
        LineTire.X1 = LineTire.X1 + (20 - vForce / 50)
        LineTourne.X2 = Cos(vCount) * 300 + 6120
        LineTourne.Y2 = Sin(vCount) * 300 + 4800
        vCount = vCount + 0.1
    Else
        TimerObj.Enabled = False
    End If
End Sub

Private Sub TimerTitre_Timer()
    LblTitre.ForeColor = "&HFF" + Hex(Int(Rnd * 5000))
End Sub

Private Sub TxtVal_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub
