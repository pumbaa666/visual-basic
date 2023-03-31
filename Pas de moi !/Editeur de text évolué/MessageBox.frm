VERSION 5.00
Begin VB.Form MessageBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin PN3.Button ButtonAnnuler 
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      Caption         =   "Annuler"
      ForeColor       =   4667408
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   1920
   End
   Begin PN3.Button ButtonNo 
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      Caption         =   "Non"
      ForeColor       =   4667408
   End
   Begin PN3.Button ButtonYes 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      Caption         =   "Oui"
      ForeColor       =   4667408
   End
   Begin PN3.Button ButtonOk 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      HighLightColor3D=   16249574
      LowLightColor3D =   9203488
      Caption         =   "Ok"
      ForeColor       =   4667408
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "yro-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1590
      TabIndex        =   10
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2025
      TabIndex        =   8
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "otes III"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   7
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label LabelVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "- Beta 1.0.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label LabelTitle 
      BackStyle       =   0  'Transparent
      Caption         =   " Message - Erreur fatale durant l'exécution"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   10
      Width           =   4815
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   4680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label LabelCaption 
      BackStyle       =   0  'Transparent
      Caption         =   $"MessageBox.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
   Begin VB.Shape ShapeWait1 
      BackColor       =   &H00B6EFFE&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   240
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape ShapeWait2 
      BackColor       =   &H00B6EFFE&
      Height          =   135
      Left            =   480
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape ShapeWait3 
      BackColor       =   &H00B6EFFE&
      Height          =   135
      Left            =   720
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape Shape 
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape ShapeStyle 
      BackColor       =   &H008C6F20&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1080
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00F5F1E7&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C5D0D1&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   1080
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Constantes de position des boutons
Const TopPosition = 1800
Const FirstPosition = 1200
Const SecondPosition = 2400
Const ThirdPosition = 3600
Const OkLeft = 4920
Const OkTop = 0
Const YesLeft = 4920
Const YesTop = 480
Const NoLeft = 4920
Const NoTop = 960
Const AnnulerLeft = 4920
Const AnnulerTop = 1440

'Déclaration de liste de types pour le style de MessageBox
Enum MsgBoxButtons
    YesNo = 0
    YesNocancel = 1
    OkOnly = 2
    OkCancel = 3
End Enum
Enum MsgBoxStyle
    Information = &H857038
    Request = &HAEC1C6
    Critical = &H50E3&
End Enum

'Déclaration des réponses possibles renvoyées
Public Enum MsgBoxResponse
    Yes = 0
    No = 1
    Cancel = 2
    Ok = 3
End Enum
Dim OkClick As Boolean
Dim YesClick As Boolean
Dim NoClick As Boolean
Dim CancelClick As Boolean

Private Sub ButtonAnnuler_Click()

CancelClick = True

End Sub

Private Sub ButtonNo_Click()

NoClick = True

End Sub

Private Sub ButtonOk_Click()

OkClick = True

End Sub

Private Sub ButtonYes_Click()

YesClick = True

End Sub

Private Sub Form_Load()

LabelVersion.Caption = NumVersion

'Redimmensionnement de la feuille
Me.Height = Shape.Height
Me.Width = Shape.Width

End Sub

Private Sub Timer_Timer()

'Pour l'animation
If ShapeWait1.BackStyle = 0 And ShapeWait2.BackStyle = 0 And ShapeWait3.BackStyle = 0 Then ShapeWait1.BackStyle = 1: Exit Sub
If ShapeWait1.BackStyle = 1 Then ShapeWait1.BackStyle = 0: ShapeWait2.BackStyle = 1: Exit Sub
If ShapeWait2.BackStyle = 1 Then ShapeWait2.BackStyle = 0: ShapeWait3.BackStyle = 1: Exit Sub
If ShapeWait3.BackStyle = 1 Then ShapeWait3.BackStyle = 0

End Sub

Public Function Message(Prompt As String, Title As String, Buttons As MsgBoxButtons, Style As MsgBoxStyle, Feuille As Form) As MsgBoxResponse

'Réinitialisation de l'emplacement et de la valeur des boutons
ButtonOk.Move OkLeft, OkTop
ButtonYes.Move YesLeft, YesTop
ButtonNo.Move NoLeft, NoTop
ButtonAnnuler.Move AnnulerLeft, AnnulerTop
OkClick = False
YesClick = False
NoClick = False
CancelClick = False

'Préparation de l'interface
LabelCaption.Caption = Prompt
LabelTitle.Caption = " Message - " & Title
ShapeStyle.BackColor = Style
If Buttons = OkOnly Then ButtonOk.Move ThirdPosition, TopPosition
If Buttons = YesNo Then ButtonYes.Move SecondPosition, TopPosition: ButtonNo.Move ThirdPosition, TopPosition
If Buttons = YesNocancel Then ButtonYes.Move FirstPosition, TopPosition: ButtonNo.Move SecondPosition, TopPosition: ButtonAnnuler.Move ThirdPosition, TopPosition

'On lance l'animation, on vérrouille et on affiche
Feuille.Enabled = False
Timer.Enabled = True
Me.Show

'On tourne en attendant l'appui sur un bouton
Do While OkClick = False And YesClick = False And NoClick = False And CancelClick = False
    DoEvents
Loop

'On renvoi la réponse
If OkClick = True Then Message = Ok
If YesClick = True Then Message = Yes
If NoClick = True Then Message = No
If CancelClick = True Then Message = Cancel

'On redonne la main aux feuilles de base et on cache la MessageBox
Me.Hide
Feuille.Enabled = True
Timer.Enabled = False

End Function

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour bouger la fenêtre
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&

End Sub
