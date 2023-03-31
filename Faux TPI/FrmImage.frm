VERSION 5.00
Begin VB.Form FrmImage 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkCote 
      Caption         =   "Voir uniquement les côtés"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblY 
      Caption         =   "Label1"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label LblX 
      Caption         =   "Label1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "FrmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
