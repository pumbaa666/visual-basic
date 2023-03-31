VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form q1Res 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Resultat 
      Height          =   5610
      Left            =   1335
      TabIndex        =   3
      Top             =   600
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   9895
      _Version        =   327680
      Cols            =   3
      BackColor       =   65280
      ForeColor       =   16711680
      BackColorFixed  =   16744576
      ForeColorFixed  =   65280
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Suite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7410
      TabIndex        =   2
      Top             =   3218
      Width           =   885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   0
      Width           =   11175
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
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   6300
      Width           =   8295
   End
End
Attribute VB_Name = "q1Res"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdok_Click()
MsgBox "Vous venez de terminer Le Questionnaire. J'espère que ça vous a plu" & vbCr & "Vous pouvez si vous le voulez recommancer." & vbCr & "Bon amusement !", 0, "Le Questionnaire"
Load frmChoice
Unload Me
End Sub
Public Sub Form_Load()
Show
lTitle.Caption = "Vérification des réponses"
Resultat.Clear
Resultat.Cols = 3
Resultat.ColWidth(0) = 600
Resultat.ColWidth(1) = 2500
Resultat.ColWidth(2) = 2500
Resultat.Row = 0
Resultat.Col = 0
Resultat.ColAlignment(0) = 3
Resultat.Text = "N°"
Resultat.Col = 1
Resultat.ColAlignment(1) = 3
Resultat.Text = "Votre Réponse"
Resultat.Col = 2
Resultat.ColAlignment(2) = 3
Resultat.Text = "Bonne Réponse"
Resultat.Row = Resultat.Row + 1

For T = 0 To N - 1
Resultat.Col = 0
Resultat.Text = T + 1
Resultat.Col = 1
Nom = UCase(q1.Text1(T).Text)
Resultat.Text = Nom
Resultat.Col = 2
Nom = UCase(Reponse(T))
Resultat.Text = Nom
If UCase(q1.Text1(T).Text) = Nom Then
Bon = Bon + 1
Else
Mauv = Mauv + 1
End If
Resultat.Rows = Resultat.Rows + 1
Resultat.Row = Resultat.Row + 1
Next
Nom = ""
N = Bon + Mauv
T = (Bon / N) * 100
If T < 10 Then Nom = UCase("Trés Mauvais")
If T > 10 And T < 21 Then Nom = UCase("Mauvais")
If T > 20 And T < 31 Then Nom = UCase("Pas Bon")
If T > 30 And T < 51 Then Nom = UCase("Moyen")
If T > 50 And T < 61 Then Nom = UCase("Bien")
If T > 60 And T < 81 Then Nom = UCase("Trés Bien")
If T > 80 Then Nom = UCase("Exelent")
Label1.Caption = " Vous avez " & Bon & " Bonnes réponses et " & Mauv & " Mauvaises COMMENTAIRE " & Nom
End Sub
