VERSION 5.00
Begin VB.Form def_touches 
   BackColor       =   &H00FF8080&
   Caption         =   "Brick Game"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "def_touches.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Remarque"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   4455
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Vous pouvez utiliser toutes les touches sauf les flèches, le pavé numérique et les touches spéciales."
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton Annul 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FF8080&
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Accept 
      BackColor       =   &H00FF8080&
      Caption         =   "Accepter"
      Height          =   375
      Left            =   600
      MaskColor       =   &H00FF8080&
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame cadr_tch 
      BackColor       =   &H00FF8080&
      Caption         =   "Touches"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.TextBox touche 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   1
         Tag             =   "Gauche"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox touche 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Index           =   3
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   4
         Tag             =   "Rotation"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox touche 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Index           =   1
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "Droite"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox touche 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Index           =   4
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "Pause"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox touche 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   3
         Tag             =   "Bas"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Pause"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Rotation"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Droite"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Bas"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Gauche"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "def_touches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gauche

Private Sub Accept_Click()

On Error GoTo titi

    Open "touche.cfg" For Output As #1
    Write #1, touche(0).Text
    Write #1, touche(1).Text
    Write #1, touche(2).Text
    Write #1, touche(3).Text
    Write #1, touche(4).Text
    Close #1
titi:
    def_touches.Hide
    Maine.Compteur.Enabled = True
End Sub

Private Sub Annul_Click()
    Call defi_touche
    def_touches.Hide
    Maine.Compteur.Enabled = True
End Sub
Private Sub Form_Activate()
    
    Call defi_touche
    If touche(cpt) <> "" Then
    
    touche(0).Text = def_tch(0)
    touche(1).Text = def_tch(1)
    touche(2).Text = def_tch(2)
    touche(3).Text = def_tch(3)
    touche(4).Text = def_tch(4)
    
    Else
    
    touche(0).Text = "a"
    touche(1).Text = "d"
    touche(2).Text = "s"
    touche(3).Text = "w"
    touche(4).Text = " "
    End If
    def_touches.Show
End Sub

Public Sub defi_touche()
    On Error GoTo titi
    
    Open "touche.cfg" For Input As #1
        Input #1, def_tch(0)
        Input #1, def_tch(1)
        Input #1, def_tch(2)
        Input #1, def_tch(3)
        Input #1, def_tch(4)
    Close #1
    
    touche(0).Text = def_tch(0)
    touche(1).Text = def_tch(1)
    touche(2).Text = def_tch(2)
    touche(3).Text = def_tch(3)
    touche(4).Text = def_tch(4)
    Exit Sub
titi:
    Close #1
    touche(0).Text = "a"
    touche(1).Text = "d"
    touche(2).Text = "s"
    touche(3).Text = "w"
    touche(4).Text = " "
End Sub

