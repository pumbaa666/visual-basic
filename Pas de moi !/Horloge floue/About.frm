VERSION 5.00
Begin VB.Form About 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "A propos de l'Horloge Floue"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   4380
      Left            =   0
      Picture         =   "About.frx":0103
      Top             =   0
      Width           =   4470
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
