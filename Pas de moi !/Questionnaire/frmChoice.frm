VERSION 5.00
Begin VB.Form frmChoice 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdapp 
      Caption         =   "A propos ..."
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "GO !"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Retour"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Choice 
      BackColor       =   &H0080FFFF&
      Caption         =   "5 choix possibles"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5055
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sélectionnez l'activité désirée :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdapp_Click()
Load frmApp
End Sub
Public Sub cmdback_Click()
Load frmSplash
Unload Me
End Sub

Public Sub cmdok_Click()
frmIntroAfr.lbli.Caption = "Bienvenuie dans le Questionnaire.  Bonne chance à vous !"
If Option1 = True Then
frmIntroAfr.lTitle.Caption = "Collectionneurs"
ElseIf Option2 = True Then
frmIntroAfr.lTitle.Caption = "Capitales"
ElseIf Option3 = True Then
frmIntroAfr.lTitle.Caption = "Cult Gén"
ElseIf Option4 = True Then
frmIntroAfr.lTitle.Caption = "S Chimiques"
ElseIf Option5 = True Then
frmIntroAfr.lTitle.Caption = "Les Simpson"
End If
Load frmIntroAfr
Unload Me
End Sub
Public Sub cmdquit_Click()
If MsgBox("Etes-vous certain de vouloir quitter cette application?" & Chr(10) & "Merci de votre visite et à bientôt j'espère.", 4 + 32 + 256, "Les Questions") = vbYes Then
End
End If
End Sub
Public Sub Form_Load()
Show
lTitle.Caption = "Choix de l'activité"
Option1.Caption = "Collectionneurs"
Option2.Caption = "Capitales"
Option3.Caption = "Cult Gén"
Option4.Caption = "S Chimiques"
Option5.Caption = "Les Simpson"
End Sub
