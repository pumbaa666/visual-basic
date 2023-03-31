VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form info_partie 
   BackColor       =   &H00FF8080&
   Caption         =   "Brick Game"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "info_partie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MCI.MMControl MMControl2 
      Height          =   615
      Left            =   5400
      TabIndex        =   21
      Top             =   3360
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1085
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl1 
      Height          =   735
      Left            =   5280
      TabIndex        =   20
      Top             =   2280
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1296
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame Info 
      BackColor       =   &H00FF8080&
      Caption         =   "Tetris"
      Height          =   4575
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Fermeture 
         Caption         =   "Fermer"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Debut 
         Caption         =   "Commencer"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         Caption         =   "Choix musique"
         Height          =   1095
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
         Begin VB.OptionButton musique4 
            BackColor       =   &H00FF8080&
            Caption         =   "Musique 4"
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton musique3 
            BackColor       =   &H00FF8080&
            Caption         =   "Musique 3"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton musique2 
            BackColor       =   &H00FF8080&
            Caption         =   "Musique 2"
            Height          =   255
            Left            =   2400
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton musique1 
            BackColor       =   &H00FF8080&
            Caption         =   "Musique 1"
            Height          =   255
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame diff 
         BackColor       =   &H00FF8080&
         Caption         =   "Niveau de difficulté"
         Height          =   1335
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   3975
         Begin VB.OptionButton op_niv_9 
            BackColor       =   &H00FF8080&
            Caption         =   "9"
            Height          =   255
            Left            =   2280
            TabIndex        =   13
            Top             =   840
            Width           =   375
         End
         Begin VB.OptionButton op_niv_8 
            BackColor       =   &H00FF8080&
            Caption         =   "8"
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   840
            Width           =   375
         End
         Begin VB.OptionButton op_niv_7 
            BackColor       =   &H00FF8080&
            Caption         =   "7"
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   840
            Width           =   375
         End
         Begin VB.OptionButton op_niv_6 
            BackColor       =   &H00FF8080&
            Caption         =   "6"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   375
         End
         Begin VB.OptionButton op_niv_5 
            BackColor       =   &H00FF8080&
            Caption         =   "5"
            Height          =   255
            Left            =   2880
            TabIndex        =   9
            Top             =   480
            Width           =   375
         End
         Begin VB.OptionButton op_niv_4 
            BackColor       =   &H00FF8080&
            Caption         =   "4"
            Height          =   255
            Left            =   2280
            TabIndex        =   8
            Top             =   480
            Width           =   375
         End
         Begin VB.OptionButton op_niv_3 
            BackColor       =   &H00FF8080&
            Caption         =   "3"
            Height          =   255
            Left            =   1680
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton op_niv_2 
            BackColor       =   &H00FF8080&
            Caption         =   "2"
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   480
            Width           =   375
         End
         Begin VB.OptionButton op_niv_1 
            BackColor       =   &H00FF8080&
            Caption         =   "1"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.TextBox Nom_joueur 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Text            =   "Player_1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Nom du joueur"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "info_partie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim musique As Byte

Private Sub Form_Click()
Call lecture

End Sub

Private Sub Form_Load()
    Call lecture
    Call meilleurs_scores.Score_lire
    Call def_touches.defi_touche
End Sub
Private Sub Debut_Click()
    Call lecture
    info_partie.Hide
    
    Maine.Show
End Sub

Private Sub Fermeture_Click()
    Unload info_partie
    End
End Sub
Function lecture()

If musique1.Value = True Then
    MMControl1.FileName = LoadResData("SON1", "MIDI")
    MMControl1.Command = "close"
    MMControl1.Command = "open"
    MMControl1.Command = "play"
End If

If musique2.Value = True Then
    MMControl1.FileName = App.Path & "\2.mid"
    MMControl1.Command = "close"
    MMControl1.Command = "open"
    MMControl1.Command = "play"
End If

If musique3.Value = True Then
    MMControl1.FileName = App.Path & "\3.mid"
    MMControl1.Command = "close"
    MMControl1.Command = "open"
    MMControl1.Command = "play"
End If

If musique4.Value = True Then
    MMControl1.FileName = App.Path & "\4.mid"
    MMControl1.Command = "close"
    MMControl1.Command = "open"
    MMControl1.Command = "play"
End If

End Function

Private Sub musique1_Click()
Call lecture
End Sub

Private Sub musique2_Click()
Call lecture
End Sub

Private Sub musique3_Click()
Call lecture
End Sub

Private Sub musique4_Click()
Call lecture
End Sub
