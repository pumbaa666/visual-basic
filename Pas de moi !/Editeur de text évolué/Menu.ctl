VERSION 5.00
Begin VB.UserControl Menu 
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ScaleHeight     =   3570
   ScaleWidth      =   3885
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         Picture         =   "Menu.ctx":0000
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1ECDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Exemple"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
