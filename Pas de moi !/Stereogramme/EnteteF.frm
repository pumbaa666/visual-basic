VERSION 5.00
Begin VB.Form EnteteF 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6510
   Icon            =   "EnteteF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "EnteteFichierBmp.EFBFileType"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "nb couleurs importantes"
      Height          =   375
      Index           =   15
      Left            =   3240
      TabIndex        =   29
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "nb couleurs ds palette"
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   28
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "résol verticale pixels"
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   27
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "résol horizontale pixels"
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   26
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "taille image octets"
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   25
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "compression"
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   24
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "nb bits par pixels"
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   23
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "nb plans utilisés (svt 1)"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "hauteur image pixels"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "largeur image pixels"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "taille entête octet"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "offset image"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "réservé...à zéro"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "taille fichier octet"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "signature fichier (BM)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "EnteteF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
