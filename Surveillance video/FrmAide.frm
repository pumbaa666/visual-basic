VERSION 5.00
Begin VB.Form FrmAide 
   Caption         =   "Aide"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.Shape ShpCam 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   $"FrmAide.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   1920
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   1800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   960
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Label Label4 
      Caption         =   "Caméra"
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
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   $"FrmAide.frx":00A6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "180°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "0°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   1200
      Top             =   360
      Width           =   2415
   End
   Begin VB.Shape ShpRay 
      Height          =   2175
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label LblDegre 
      Caption         =   "90°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   2760
      Width           =   495
   End
End
Attribute VB_Name = "FrmAide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    FrmMain.Show
    FrmAide.Hide
End Sub

Private Sub Form_Activate()
    ShpCam.FillColor = FrmMain.ShpCam.FillColor
    ShpCam.BorderColor = FrmMain.ShpCam.BorderColor
End Sub
