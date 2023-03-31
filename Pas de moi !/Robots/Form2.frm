VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programmation des robots"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Enlever des points d'energie même si l'action est impossible"
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   5880
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Afficher l'historique à la fin du match"
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   5520
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CommandButton BtSave 
      Caption         =   "Sauvegarder le programme de ce robot"
      Height          =   375
      Left            =   2760
      TabIndex        =   70
      Top             =   4320
      Width           =   3975
   End
   Begin VB.VScrollBar VS 
      Height          =   285
      Left            =   6600
      Max             =   2
      Min             =   4
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5160
      Value           =   4
      Width           =   255
   End
   Begin VB.TextBox Tobs 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   65
      Text            =   "50"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Tdist 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5760
      TabIndex        =   63
      Text            =   "4"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Ttimer 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      TabIndex        =   61
      Text            =   "200"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Tnrj 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      TabIndex        =   40
      Text            =   "1500"
      Top             =   4785
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Commencer"
      Height          =   375
      Left            =   5160
      TabIndex        =   39
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Inst 
      Caption         =   "TV"
      Height          =   375
      Index           =   9
      Left            =   2760
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "TH"
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "TT_"
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "IN"
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "MI"
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "FT"
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "PS"
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "AL"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Inst 
      Caption         =   "DD_"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Instructions de programmation"
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2535
      Begin VB.Label LbCout 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         MouseIcon       =   "Form2.frx":0000
         TabIndex        =   29
         ToolTipText     =   "Instruction de secours"
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         MouseIcon       =   "Form2.frx":030A
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         MouseIcon       =   "Form2.frx":0614
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         MouseIcon       =   "Form2.frx":091E
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         MouseIcon       =   "Form2.frx":0C28
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   240
         MouseIcon       =   "Form2.frx":0F32
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   240
         MouseIcon       =   "Form2.frx":123C
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   240
         MouseIcon       =   "Form2.frx":1546
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   240
         MouseIcon       =   "Form2.frx":1850
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   240
         MouseIcon       =   "Form2.frx":1B5A
         TabIndex        =   20
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   240
         MouseIcon       =   "Form2.frx":1E64
         TabIndex        =   19
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   1320
         MouseIcon       =   "Form2.frx":216E
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1320
         MouseIcon       =   "Form2.frx":2478
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1320
         MouseIcon       =   "Form2.frx":2782
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1320
         MouseIcon       =   "Form2.frx":2A8C
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   1320
         MouseIcon       =   "Form2.frx":2D96
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   1320
         MouseIcon       =   "Form2.frx":30A0
         TabIndex        =   13
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   1320
         MouseIcon       =   "Form2.frx":33AA
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1320
         MouseIcon       =   "Form2.frx":36B4
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1320
         MouseIcon       =   "Form2.frx":39BE
         TabIndex        =   10
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Bloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   1320
         MouseIcon       =   "Form2.frx":3CC8
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre de joueurs :"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   69
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label LbNbr 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   68
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Obstacles à générer :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   66
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Portée du radar :"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   64
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Durée d'un tour (ms) :"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   62
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Lb 
      Caption         =   "Tir vertical"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   60
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Tir horizontal"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   59
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Test de proximité"
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   58
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Invisibilité"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   57
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Pose d'une mine"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   56
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Fuite"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   55
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Poursuite"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   54
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Déplacement aléatoire"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   53
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Lb 
      Caption         =   "Déplacement choisi"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   52
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Coût"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   51
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "5"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   50
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   49
      Top             =   1440
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   48
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   47
      Top             =   2160
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   46
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "20"
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   45
      Top             =   2880
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   44
      Top             =   3240
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   43
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Index           =   9
      Left            =   5160
      TabIndex        =   42
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label Label3 
      Caption         =   "Energie :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Robot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   5160
      MouseIcon       =   "Form2.frx":3FD2
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Robot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3480
      MouseIcon       =   "Form2.frx":42DC
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Robot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1800
      MouseIcon       =   "Form2.frx":45E6
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Robot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   120
      MouseIcon       =   "Form2.frx":48F0
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "- 20"
      Height          =   255
      Index           =   12
      Left            =   5640
      TabIndex        =   3
      Top             =   3960
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "- 20"
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   2
      Top             =   3600
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "- 200"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   1
      Top             =   2520
      Width           =   450
   End
   Begin VB.Label Label2 
      Caption         =   "Dégats"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bloc_Click(Index As Integer)
Dim I As Long

For I = 0 To UBound(Pas, 2)
Bloc(I).BackColor = vbWhite
Next I
Bloc(Index).BackColor = 16773103

BtSave_Click

End Sub

Private Sub BtSave_Click()
Dim I, J As Long

For I = 1 To NbR
    If Robot(I).BackColor = 16773103 Then
        For J = 0 To UBound(Pas, 2)
        Pas(I, J) = Bloc(J).Caption
        Next J
    End If
Next I

For I = 1 To NbR
If Robot(I).BackColor = 16773103 Then Robot_Click (I)
Next I

End Sub

Private Sub Command1_Click()
Dim I, J As Long

If Val(Tnrj) < 300 Or Val(Tnrj) > 10000 Then Exit Sub
If Val(Ttimer) < 1 Or Val(Ttimer) > 30000 Then Exit Sub
If Val(Tdist) < 1 Or Val(Tdist) > 50 Then Exit Sub
If Val(Tobs) < 1 Or Val(Tobs) > 400 Then Exit Sub
'verifie si toutes les instructions sont remplies
For I = 1 To NbR
    For J = 0 To UBound(Pas, 2)
        If Pas(I, J) = "" Then Exit Sub
    Next J
Next I

NbR = LbNbr
Form1.Timer1.Interval = Val(Ttimer)
DisRep = Tdist
NbreBlocs = Tobs

For I = 1 To NbR
    Nom(I) = Robot(I)
    Robot_Click (I)
    PV(I) = Tnrj
Next I

For I = 1 To NbR
    Robot_Click (I)
    Open App.Path & "\programmes\" & I & ".txt" For Output As #1
        For J = 0 To UBound(Pas, 2)
            Print #1, Pas(I, J)
        Next J
    Close #1
Next I

PtEnMoinsSiPasPossible = Check2.Value


AffichageCteRendu = Check1.Value

Unload Form2

End Sub

Private Sub Form_Load()
Dim I As Long
Dim K As Integer
Dim Instruction As String

LbNbr = NbR
'met les pas de programmes en mémoire
For I = 1 To NbR
    Robot(I).Caption = Nom(I)
    Robot(I).Visible = True
    K = 0
    Open App.Path & "\programmes\" & I & ".txt" For Input As #1
        While Not EOF(1)
            Line Input #1, Instruction
            Pas(I, K) = Instruction
            K = K + 1
        Wend
    Close #1
Next I

For I = 0 To UBound(Pas, 2)
Bloc(I).Visible = True
If I <> 0 Then Bloc(I).ToolTipText = "Instruction n° " & I
Next I

Robot(1).BackColor = 16773103
Bloc(0).BackColor = 16773103

Robot_Click (1)
End Sub

Private Sub Inst_Click(Index As Integer)
Dim I As Long

If Inst(Index).Caption = "DD_" And Bloc(0).BackColor <> 16773103 Then
    Form3.Show 1
    Exit Sub
End If

For I = 0 To UBound(Pas, 2)
    If Bloc(I).BackColor = 16773103 And Inst(Index).Caption = "TH" And Len(Bloc(I).Caption) < 3 Then Exit Sub
    If Bloc(I).BackColor = 16773103 And Inst(Index).Caption = "TV" And Len(Bloc(I).Caption) < 3 Then Exit Sub
    If Bloc(I).BackColor = 16773103 And Bloc(I).Caption = "TT_" And Inst(Index).Caption <> "DD_" Then
        Bloc(I).Caption = "TT_" & Inst(Index).Caption & "_"
        Exit Sub
    End If
    If Bloc(I).BackColor = 16773103 And Len(Bloc(I).Caption) = 6 And Inst(Index).Caption <> "DD_" Then
        Bloc(I).Caption = Bloc(I).Caption & Inst(Index).Caption
        Exit Sub
    End If
Next I

If Bloc(0).BackColor = 16773103 And Inst(Index).Caption = "DD_" Then
    MsgBox "Interdit en instruction de secours."
    Exit Sub
End If

If Bloc(0).BackColor = 16773103 And Inst(Index).Caption = "TT_" Then
    MsgBox "Interdit en instruction de secours."
    Exit Sub
End If


For I = 0 To UBound(Pas, 2)
If Bloc(I).BackColor = 16773103 Then Bloc(I).Caption = Inst(Index).Caption
Next I

End Sub

Private Sub LbNbr_Change()
Dim I As Long

NbR = LbNbr

For I = 1 To 4: Robot(I).Visible = False: Next I
For I = 1 To NbR: Robot(I).Visible = True: Next I

Robot_Click (1)
End Sub

Private Sub Robot_Click(Index As Integer)
Dim J, I As Long
Dim Cmin, Cmax As Integer
Dim C1, C2 As String

For J = 0 To UBound(Pas, 2): Bloc(J).Caption = Pas(Index, J): Next J

For I = 1 To NbR: Robot(I).BackColor = vbWhite: Next I

Robot(Index).BackColor = 16773103

'calcul du coût par tour
For I = 1 To UBound(Pas, 2) 'on ne compte pas le pas de secours
If Left$(Bloc(I).Caption, 2) = "DD" Then Cmin = Cmin + 5: Cmax = Cmax + 5
If Left$(Bloc(I).Caption, 2) = "AL" Then Cmin = Cmin + 1: Cmax = Cmax + 1
If Left$(Bloc(I).Caption, 2) = "PS" Then Cmin = Cmin + 4: Cmax = Cmax + 4
If Left$(Bloc(I).Caption, 2) = "FT" Then Cmin = Cmin + 4: Cmax = Cmax + 4
If Left$(Bloc(I).Caption, 2) = "MI" Then Cmin = Cmin + 10: Cmax = Cmax + 10
If Left$(Bloc(I).Caption, 2) = "IN" Then Cmin = Cmin + 20: Cmax = Cmax + 20
If Left$(Bloc(I).Caption, 2) = "TT" Then
    Cmin = Cmin + 4: Cmax = Cmax + 4
    C1 = Mid$(Bloc(I).Caption, 4, 2)
    C2 = Mid$(Bloc(I).Caption, 7, 2)
    If C1 = "AL" Then C1 = 1
    If C1 = "PS" Then C1 = 4
    If C1 = "FT" Then C1 = 4
    If C1 = "MI" Then C1 = 10
    If C1 = "IN" Then C1 = 20
    If C1 = "TH" Then C1 = 3
    If C1 = "TV" Then C1 = 3
    If C2 = "AL" Then C2 = 1
    If C2 = "PS" Then C2 = 4
    If C2 = "FT" Then C2 = 4
    If C2 = "MI" Then C2 = 10
    If C2 = "IN" Then C2 = 20
    If C2 = "TH" Then C2 = 3
    If C2 = "TV" Then C2 = 3
    If C1 = C2 Then Cmin = Cmin + C1: Cmax = Cmax + C1
    If C1 < C2 Then Cmin = Cmin + C1: Cmax = Cmax + C2
    If C1 > C2 Then Cmin = Cmin + C2: Cmax = Cmax + C1
End If
Next I

LbCout = "Coût par tour : " & Cmin & " à " & Cmax
End Sub

Private Sub Robot_DblClick(Index As Integer)
Dim V

10 V = InputBox("Entrez le nouveau nom du robot n° " & Index)
If V = "" Then GoTo 10
Nom(Index) = V
Robot(Index) = V
End Sub

Private Sub VS_Change()
LbNbr = VS.Value
End Sub
