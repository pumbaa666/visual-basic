VERSION 5.00
Begin VB.Form FrmOption 
   Caption         =   "Options de connexion"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordinateur distant"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox CmbRemoteHost 
         Height          =   315
         ItemData        =   "FrmOption.frx":0000
         Left            =   240
         List            =   "FrmOption.frx":0013
         TabIndex        =   6
         Text            =   "Nom ou IP"
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox CmbRemotePort 
         Height          =   315
         ItemData        =   "FrmOption.frx":0080
         Left            =   240
         List            =   "FrmOption.frx":0090
         TabIndex        =   5
         Text            =   "1001"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "N° de port"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label LblIp 
         Caption         =   "IP Distant :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordinateur local"
      Height          =   2055
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox TxtNom 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Votre nom"
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox CmbLocalPort 
         Height          =   315
         ItemData        =   "FrmOption.frx":00AC
         Left            =   240
         List            =   "FrmOption.frx":00BC
         TabIndex        =   1
         Text            =   "1002"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "N° de port"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label LblLocalIP 
         Caption         =   "Votre IP : "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    FrmMain.Show
    FrmOption.Hide
End Sub

Private Sub CmbLocalPort_Click()
    If CmbRemotePort.Text = CmbLocalPort.Text Then
        MsgBox "Vous ne pouvez pas définir le port distant égal au port local", vbCritical, "Erreur"
        CmbLocalPort.Text = "N° de port"
    End If
End Sub

Private Sub CmbLocalPort_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub

Private Sub CmbRemotePort_Click()
    If CmbRemotePort.Text = CmbLocalPort.Text Then
        MsgBox "Vous ne pouvez pas définir le port distant égal au port local", vbCritical, "Erreur"
        CmbRemotePort.Text = "N° de port"
    End If
End Sub

Private Sub CmbRemotePort_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub

Private Sub CmdOptions_Click()
    FrmOption.Show
    FrmMain.Hide
End Sub

Private Sub Form_Load()
    LblLocalIP.Caption = "Votre IP : " & FrmMain.WSTCP.LocalIP
End Sub
