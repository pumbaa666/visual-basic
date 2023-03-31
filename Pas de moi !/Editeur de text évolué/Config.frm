VERSION 5.00
Begin VB.Form Config 
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8D0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   3615
      Begin VB.OptionButton OptionPN3RTF 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8D0C0&
         Caption         =   "Assigner les fichiers à Pyro-Notes III"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3495
      End
      Begin VB.OptionButton OptionWordpadRTF 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8D0C0&
         Caption         =   "Assigner les fichiers à Wordpad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8D0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3615
      Begin VB.OptionButton OptionNotepadTXT 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8D0C0&
         Caption         =   "Assigner les fichiers à Notepad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3015
      End
      Begin VB.OptionButton OptionPN3TXT 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8D0C0&
         Caption         =   "Assigner les fichiers à Pyro-Notes III"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
   End
   Begin PN3.Button ButtonCancel 
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   192
      BackColor       =   14938367
      Caption         =   "Annuler"
      ForeColor       =   192
   End
   Begin PN3.Button ButtonOk 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   6050618
      BackColor       =   15526369
      Caption         =   "Ok"
   End
   Begin PN3.Button ButtonApply 
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   6050618
      BackColor       =   15526369
      Caption         =   "Appliquer"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8D0C0&
      Caption         =   " Prise en charge du type RTF "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   120
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8D0C0&
      Caption         =   " Prise en charge du type TXT "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   2550
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   120
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label LabelTitle 
      BackStyle       =   0  'Transparent
      Caption         =   " Configuration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   15
      Width           =   4095
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00D8D0C0&
      BackStyle       =   1  'Opaque
      Height          =   2895
      Left            =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00B4A587&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonApply_Click()

If OptionNotepadTXT.Value = True Then
    AssignNotepadTXT
Else
    AssignPN3TXT
End If
If OptionWordpadRTF.Value = True Then
    AssignWordpadRTF
Else
    AssignPN3RTF
End If

End Sub

Private Sub ButtonCancel_Click()

Me.Hide
Main.Enabled = True

End Sub

Private Sub ButtonOk_Click()

ButtonApply_Click
ButtonCancel_Click

End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour bouger la fenêtre
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

