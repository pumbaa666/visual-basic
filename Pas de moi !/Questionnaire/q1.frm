VERSION 5.00
Begin VB.Form q1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdback 
      Caption         =   "Précédent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   113
      TabIndex        =   9
      Top             =   3870
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Suivant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5003
      TabIndex        =   8
      Top             =   3870
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.Label lTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1658
      TabIndex        =   10
      Top             =   3690
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2670
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6225
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "q1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub cmdback_Click()
T = T - 1
If T < 0 Then
Unload Me
Load frmmenu
Exit Sub
End If
If T = 0 Then
lTitle.Caption = "Question n°1"
Label5.Caption = "Vous avez donné comme réponse " & UCase(Text1(0).Text)
Else
Label5.Caption = "Vous avez donné comme réponse " & UCase(Text1(T).Text)
End If
Select Case T
Case 1
lTitle.Caption = "Question n°2"
Case 2
lTitle.Caption = "Question n°3"
Case 3
lTitle.Caption = "Question n°4"
Case 4
lTitle.Caption = "Question n°5"
Case 5
lTitle.Caption = "Question n°6"
Case 6
lTitle.Caption = "Question n°7"
Case 7
lTitle.Caption = "Question n°8"
Case 8
lTitle.Caption = "Question n°9"
Case 9
lTitle.Caption = "Question n°10"
Case 10
lTitle.Caption = "Question n°11"
End Select
Combo1.Text = ""
End Sub
Public Sub cmdok_Click()
Label5.Caption = ""
Text1(T).Text = LCase(Combo1.Text)
T = T + 1
Combo1.Text = ""
If T < N - 1 Then
lTitle.Caption = "Question n°" & T + 1
Label3.Caption = Question(T)
End If
If T = N - 1 Then
lTitle.Caption = "Question n°" & T + 1
Label3.Caption = Question(T)
cmdOK.Caption = "Réponse !"
End If
If T = N Then
Load q1Res
Unload q1
End If
End Sub
Public Sub Form_Load()
Show
lTitle.Caption = "Question 1"
Label1.Caption = "Tapez la réponse de la question dans l'espace réservé à cet effet."
Label2.Caption = "Question"
Label4.Caption = "Réponse"
Text1(0).Visible = False
N = 0
If frmIntroAfr.lTitle.Caption = "Collectionneurs" Then
FicheReponse = "RéponseUn.txt"
FicheQuestion = "QuestionUn.txt"
ElseIf frmIntroAfr.lTitle.Caption = "Capitales" Then
FicheReponse = "RéponseDeux.txt"
FicheQuestion = "QuestionDeux.txt"
ElseIf frmIntroAfr.lTitle.Caption = "Cult Gén" Then
FicheReponse = "RéponseTrois.txt"
FicheQuestion = "QuestionTrois.txt"
ElseIf frmIntroAfr.lTitle.Caption = "S Chimiques" Then
FicheReponse = "RéponseQuatre.txt"
FicheQuestion = "QuestionQuatre.txt"
ElseIf frmIntroAfr.lTitle.Caption = "Les Simpson" Then
FicheReponse = "RéponseCinq.txt"
FicheQuestion = "QuestionCinq.txt"
End If
Open FicheReponse For Input As #1
While Not EOF(1)
Line Input #1, Nom
If Trim(Nom) <> "" Then
Combo1.AddItem Right(UCase(Nom), Len(Nom) - 3)
N = N + 1
End If
Wend
Close
For T = 1 To N
Load Text1(T)
Next
ReDim Question(N), Reponse(N)
Open FicheQuestion For Input As #1
For T = 0 To N - 1
Line Input #1, Nom
If Trim(Nom) <> "" Then
Question(T) = Right(UCase(Nom), Len(Nom) - 3)
End If
Next
Close
Label3.Caption = Question(0)
Open FicheReponse For Input As #1
For T = 0 To N - 1
Line Input #1, Nom
If Trim(Nom) <> "" Then
M = Val(Left(Nom, 3)) - 1
Reponse(M) = Right(UCase(Nom), Len(Nom) - 3)
End If
Next
Close
T = 0
If N = 0 Then
MsgBox Chr(34) & frmIntroAfr.lTitle.Caption & Chr(34) & "  Fichier VIDE"
Load frmcharg
End If
End Sub

