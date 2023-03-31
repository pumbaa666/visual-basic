VERSION 5.00
Begin VB.Form frmApp 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2298.425
   ScaleMode       =   0  'User
   ScaleWidth      =   5310.337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5655
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      Begin VB.Label lTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   480
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.569
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ceci est un petit quiz sur des Questions créer en Visual Basic 5.0 par Sébastian DANCOT.Modifier Serge Cheval"
      ForeColor       =   &H8000000D&
      Height          =   570
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   4365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.483
      Y1              =   1573.696
      Y2              =   1573.696
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   3870
   End
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub cmdok_Click()
Unload Me
End Sub
Public Sub Form_Load()
Show
Me.Caption = "About " & App.Title
lTitle.Caption = "A propos ..."
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblTitle.Caption = "A propos ..."
End Sub
