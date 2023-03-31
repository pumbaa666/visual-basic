VERSION 5.00
Begin VB.Form FrmImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox ChkCote 
      Caption         =   "Voir uniquement les côtés"
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label LblY 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label LblX 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "FrmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FrmImage.Top = FrmMain.Top
    FrmImage.Left = FrmMain.Left + 500 + FrmMain.Height
    FrmImage.Height = FrmMain.Height
    FrmImage.Width = FrmMain.Width
End Sub
