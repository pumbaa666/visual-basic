VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "lumière arrêt défil"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "lumière majuscule"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "lumière numlock"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simuler apuie touche screenshot"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C


Private Sub Command1_Click()
keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0

End Sub

Private Sub Command2_Click()
    keybd_event VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_NUMLOCK, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0

End Sub

Private Sub Command3_Click()
    keybd_event VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0

End Sub

Private Sub Command4_Click()
    keybd_event VK_SCROLL, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SCROLL, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0

End Sub
