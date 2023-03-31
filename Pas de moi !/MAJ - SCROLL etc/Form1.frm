VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clavier Tray"
   ClientHeight    =   255
   ClientLeft      =   17265
   ClientTop       =   14040
   ClientWidth     =   1845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   1845
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   1320
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   582
            MinWidth        =   2
            TextSave        =   "Maj"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   714
            MinWidth        =   2
            TextSave        =   "Num"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   2
            TextSave        =   "INSER"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   2
            TextSave        =   "DÉFIL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   15
      Left            =   4200
      Picture         =   "Form1.frx":0442
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   14
      Left            =   3600
      Picture         =   "Form1.frx":0767
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   13
      Left            =   3000
      Picture         =   "Form1.frx":0A8C
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   2400
      Picture         =   "Form1.frx":0DB1
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   1800
      Picture         =   "Form1.frx":10D6
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   1200
      Picture         =   "Form1.frx":13FB
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   600
      Picture         =   "Form1.frx":1720
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   0
      Picture         =   "Form1.frx":1A45
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   4200
      Picture         =   "Form1.frx":1D6A
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   3600
      Picture         =   "Form1.frx":208F
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   3000
      Picture         =   "Form1.frx":23B4
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   2400
      Picture         =   "Form1.frx":26D9
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   1800
      Picture         =   "Form1.frx":29FE
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "Form1.frx":2D23
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":3048
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":336D
      Top             =   0
      Width           =   480
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Cache 
         Caption         =   "&Cacher"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu S0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Quitter 
         Caption         =   "&Quitter"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&A propos"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type IconeTray
    cbSize As Long      'Taille de l'icône (en octets)
    hwnd As Long        'Handle de la fenêtre chargée de recevoir les messages envoyés lors des évènements sur l'icône (clics, doubles-clics...)
    uID As Long         'Identificateur de l'icône
    uFlags As Long
    uCallbackMessage As Long    'Messages à renvoyer
    hIcon As Long               'Handle de l'icône
    szTip As String * 64        'Texte à mettre dans la bulle d'aide
End Type
Dim IconeT As IconeTray


'Constantes nécessaires
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = &O2

Private Const AJOUT = &H0
Private Const MODIF = &H1
Private Const SUPPRIME = &H2
Private Const MOUSEMOVE = &H200
Private Const MESSAGE = &H1
Private Const Icone = &H2
Private Const TIP = &H4

Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205

'API nécessaire
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconeTray) As Boolean
'Etat des touches
Private Declare Function GetKeyState Lib "user32" (ByVal iVirtualKey As Integer) As Long
Private Sub Cache_Click()
If Cache.Checked Then
    Cache.Checked = False
    Backward Me
    Me.Hide
Else
    Cache.Checked = True
    Forward Me
End If
End Sub

Private Sub Form_Load()

'************************************************************
'* NOM : System Tray
'* DATE : 20/06/1997
'*
'* AUTEUR : Antoine de Montgolfier ( Antoine@vbasic.org )
'*
'* CODE TROUVE SUR "Le petit monde de Visual Basic"
'*                 http://www.vbasic.org
'*
'* DESCRIPTION :
'* Cet exemple vous montre comment placer des icônes
'* dans le système tray (la partie à droite de la barre des
'* tâches où se trouve, entre autres, l'horloge), comment
'* y mettre un menu, et comment traiter les évènements que
'* l'utilisateur crée (ex : double clic, clic du bouton
'* droit...). Il vous est aussi possible de créer une bulle
'* d'aide lorsque la souris reste sur l'icône.
'*
'************************************************************

'Positionne la forme en bas à droite
Left = Screen.Width - Width - 50
Top = Screen.Height - Height - 950
'Préparation de la variable IconeT
IconeT.cbSize = Len(IconeT)             'Taille de l'icône en octet
IconeT.hwnd = Me.hwnd                   'Handle de l'application (pour qu'elle reçoive les messages envoyés lors d'un clic, double-clic...
IconeT.uID = 1&                         'Identificateur de l'icône
IconeT.uFlags = Icone Or TIP Or MESSAGE
IconeT.uCallbackMessage = MOUSEMOVE     'Renvoyer les messages concernant l'action de la souris
IconeT.hIcon = Image1(0).Picture           'Mettre en icône l'image qui est dans le contrôle "Image1"
IconeT.szTip = "Clavier Tray" & Chr$(0) 'Texte de la bulle d'aide

'Appel de la fonction pour mettre l'icône dans le système tray
Shell_NotifyIcon AJOUT, IconeT


Me.Hide     'Cache la fenêtre
App.TaskVisible = False     'Retire le bouton de l'application de la barre
                            'des tâches
'Cache_Click
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static rec As Boolean, Msg As Long
Dim Indicat$

'Se produit lorsque l'utilisateur agit avec la souris sur
'l'icône placée dans le système tray
Indicat$ = ""
If LireEtatMaj > 0 Then Indicat$ = "Maj"
If LireEtatNum > 0 Then
    If Indicat$ <> "" Then Indicat$ = Indicat$ + " "
    Indicat$ = Indicat$ + "Num"
End If
If LireEtatInsert > 0 Then
    If Indicat$ <> "" Then Indicat$ = Indicat$ + " "
    Indicat$ = Indicat$ + "Insert"
End If

If LireEtatScroll > 0 Then
    If Indicat$ <> "" Then Indicat$ = Indicat$ + " "
    Indicat$ = Indicat$ + "Defil"
End If

IconeT.szTip = Indicat$ & Chr$(0) 'Texte de la bulle d'aide
Shell_NotifyIcon MODIF, IconeT

Msg = x / Screen.TwipsPerPixelX
'If rec = False Then
'    rec = True
    Select Case Msg     'Différentes possibilité d'action
        Case DOUBLE_CLICK_GAUCHE:   'mettez
                                    'ici
        Case BOUTON_GAUCHE_POUSSE:  'ce
        Case BOUTON_GAUCHE_LEVE:    'que
            About_Click
        Case DOUBLE_CLICK_DROIT:    'vous
        Case BOUTON_DROIT_POUSSE:   'voudrez
        Case BOUTON_DROIT_LEVE:     'qu'il se passe
            PopupMenu Menu, , , , About     'fait apparaitre le menu
            '"A propos de" apparaitra en gras
    End Select
'    rec = False
'End If

End Sub

Private Sub Quitter_Click()

Unload Me   'retire la fenêtre
End

End Sub

Private Sub Timer1_Timer()
Dim Indeximg
Indeximg = LireEtatMaj + LireEtatNum + LireEtatScroll + LireEtatInsert
'Me.Icon = Image1(Indeximg).Picture
IconeT.hIcon = Image1(Indeximg).Picture            'Mettre en icône l'image qui est dans le contrôle "Image1"
Shell_NotifyIcon MODIF, IconeT
End Sub
Private Sub About_Click()
 frmAbout.Show
End Sub

