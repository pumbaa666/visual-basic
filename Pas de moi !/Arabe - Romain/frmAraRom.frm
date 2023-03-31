VERSION 5.00
Begin VB.Form frmAraRom 
   BackColor       =   &H00D0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AraRom"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmAraRom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   57
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConversion 
      BackColor       =   &H00D0FFFF&
      Caption         =   "&Conversion"
      Default         =   -1  'True
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtRomain 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtArabe 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTexte 
      BackStyle       =   0  'Transparent
      Caption         =   "Chiffre romain :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label lblTexte 
      BackStyle       =   0  'Transparent
      Caption         =   "Chiffre arabe :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "frmAraRom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConversion_Click()

On Error GoTo Erreur
txtRomain.Text = ArabeRomain(txtArabe.Text)
Exit Sub

Erreur:
MsgBox "La valeur doit être comprise entre 1 et 3999 inclus.", vbCritical + vbOKOnly, "Erreur"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Unload Me
End

End Sub
