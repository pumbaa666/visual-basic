VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Avancement 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Avancement"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   Icon            =   "Avancement3D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar Pgbar 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   61
   End
   Begin VB.Label phase1 
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Label phase1 
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label phase1 
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label lbl 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Avancement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
