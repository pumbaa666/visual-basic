VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pente"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk 
      Caption         =   "Conserver les valeurs"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton CmdZero 
      Caption         =   "&Remise à 0"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Timer ClkTombe 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   4440
   End
   Begin VB.Timer ClkPente 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   5160
   End
   Begin VB.CommandButton CmdTest 
      Caption         =   "&Tester"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox TxtVal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Text            =   "1"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox TxtVal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   4440
      TabIndex        =   1
      Text            =   "90"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label LblTombe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line LineObj 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   20
      X1              =   3480
      X2              =   3960
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line LinePente 
      BorderWidth     =   5
      X1              =   1560
      X2              =   4200
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label1 
      Caption         =   "Angle d'inclinaison (0 à 90°)"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Coefficient de frottement (0 à 1)"
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
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCount As Double
Dim vLngObj As Integer
Dim vHautObj As Integer
Const Angle = 0
Const Mu = 1

Private Sub ClkTombe_Timer()
Static vX1 As Integer
Static vX2 As Integer
Static vCount As Integer
    If LineObj.Y1 < 5400 Then
        LineObj.Y1 = LineObj.Y1 + (LineObj.Y2 - LineObj.Y1)
        LineObj.Y2 = LineObj.Y2 + (LineObj.Y2 - LineObj.Y1)
        LineObj.X1 = LineObj.X1 - (LineObj.X2 - LineObj.X1)
        LineObj.X2 = LineObj.X2 - (LineObj.X2 - LineObj.X1)
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdTest_Click()
    If TxtVal(Angle) = "" Or TxtVal(Mu) = "" Then
        MsgBox "Il manque des paramètres!!!", vbCritical, "Pente"
    Else
        If Int(TxtVal(Angle)) > 90 Then
            MsgBox "L'angle doit être compris entre 0 et 90 °", vbCritical, "Pente"
        ElseIf Int(TxtVal(Mu)) > 1 Then
            MsgBox "Le coefficient de frottement doit être compris entre 0 et 1", vbCritical, "Pente"
        Else
            ClkPente.Enabled = True
            CmdTest.Enabled = False
        End If
    End If
End Sub

Private Sub ClkPente_Timer()
Static vLngPente As Integer
Static vVirX As Double
Static vVirY As Double
Dim vTiens As Double
    If vCount = 0 Then
        vCount = 3.2
        vLngPente = LinePente.X2 - LinePente.X1
    End If
    LinePente.X2 = -Cos(vCount) * vLngPente + LinePente.X1
    LinePente.Y2 = Sin(vCount) * vLngPente + LinePente.Y1
    
    vVirX = -Cos(vCount + 90 * 3.141592654 / 180) * 120 + LinePente.X1
    vVirY = Sin(vCount + 90 * 3.141592654 / 180) * 120 + LinePente.Y1
    
    LineObj.X1 = -Cos(vCount) * (vLngPente - 800) + vVirX
    LineObj.Y1 = Sin(vCount) * (vLngPente - 800) + vVirY
    
    LineObj.X2 = -Cos(vCount) * vLngObj + LineObj.X1
    LineObj.Y2 = Sin(vCount) * vLngObj + LineObj.Y1
    vCount = vCount + 0.01
    If vCount - 3.148 >= Int(TxtVal(Angle).Text) * 3.14159265358979 / 180 Then
        ClkPente.Enabled = False
        LblTombe.ForeColor = &HFF0000
        LblTombe.Caption = "Ca tiens!"
    End If
    If CDbl(TxtVal(Mu)) * Cos(vCount) > Sin(vCount) Then
        ClkPente.Enabled = False
        ClkTombe.Enabled = True
        LblTombe.ForeColor = &HFF&
        LblTombe.Caption = "BOUM!"
    End If
End Sub

Private Sub CmdZero_Click()
    If Chk.Value <> Checked Then
        TxtVal(Mu).Text = ""
        TxtVal(Angle).Text = ""
    End If
    LinePente.Y2 = LinePente.Y1
    LinePente.X2 = 4220
    ClkPente.Enabled = False
    ClkTombe.Enabled = False
    CmdTest.Enabled = True
    LineObj.X1 = 3480
    LineObj.X2 = 3960
    LineObj.Y1 = 5280
    LineObj.Y2 = 5280
    LblTombe.Caption = ""
    vCount = 0
End Sub

Private Sub Form_Load()
    vLngObj = LineObj.X2 - LineObj.X1
    vHautObj = LineObj.Y2 - LineObj.Y1
End Sub

Private Sub TxtVal_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub
