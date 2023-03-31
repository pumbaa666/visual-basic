VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   3195
   ClientLeft      =   2355
   ClientTop       =   2235
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Ce logiciel permet de voir dans le systray les indicateurs : Maj, Num, Insert et Scroll"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "avec Microsoft Visual Basic 6.0"
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   765
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "amelie_leitz@hotmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Créé par Amélie Leitz en Octobre 2003 "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   555
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Clavier Tray"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Indeximg
Indeximg = LireEtatMaj + LireEtatNum + LireEtatScroll + LireEtatInsert
'Me.Icon = Image1(Indeximg).Picture
Image1.Picture = Form1.Image1(Indeximg).Picture            'Mettre en icône l'image qui est dans le contrôle "Image1"

End Sub
