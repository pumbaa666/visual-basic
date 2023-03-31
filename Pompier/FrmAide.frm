VERSION 5.00
Begin VB.Form FrmAide 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aide"
   ClientHeight    =   4275
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Cliquez ensuite sur Ok pour valider votre choix, ou sur Annuler, pour annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label Label4 
      Caption         =   "Si une liste s'affiche, cliquez sur un des outils pour changer son nom, sa couleur ou son emplacement"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Ou ""Tout ceux qui se trouve en Réparation"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Par exemple ""Tout les bleu"""
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Choisissez quels outils afficher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "FrmAide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    FrmAide.Hide
End Sub
