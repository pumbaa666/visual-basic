VERSION 5.00
Begin VB.Form FrmCouleur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Couleur"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LblCouleur 
      Caption         =   "Noir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   240
      MouseIcon       =   "FrmCouleur.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LblCouleur 
      Caption         =   "Violet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   3
      Left            =   1440
      MouseIcon       =   "FrmCouleur.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LblCouleur 
      Caption         =   "Vert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   1440
      MouseIcon       =   "FrmCouleur.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label LblCouleur 
      Caption         =   "Rouge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   240
      MouseIcon       =   "FrmCouleur.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LblCouleur 
      Caption         =   "Bleu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   240
      MouseIcon       =   "FrmCouleur.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmCouleur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LblCouleur_Click(Index As Integer)
    vCouleur = LblCouleur(Index).Caption
    
    If FrmMain.OptionPartieSolo.Checked = False Then
        If vQui = "Client" Then
            FrmMain.ShpTerrain((DimX + 1) * (DimY + 1) - 1).Picture = LoadPicture("images/" & vCouleur & "/horiz.bmp")
        Else
            FrmMain.ShpTerrain(0).Picture = LoadPicture("images/" & vCouleur & "/horiz.bmp")
        End If
        FrmOptMulti.Wsk.SendData "[COULEUR]" & vCouleur
    Else
         FrmMain.ShpTerrain(0).Picture = LoadPicture("images/" & vCouleur & "/horiz.bmp")
    End If
    
    FrmCouleur.Hide
End Sub
