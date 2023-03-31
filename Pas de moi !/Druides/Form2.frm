VERSION 5.00
Begin VB.Form Manger 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Manger"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Poissons"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4455
      Begin VB.ListBox List3 
         Height          =   450
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Non"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Oui"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Qté :"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fruits"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
      Begin VB.ListBox List2 
         Height          =   450
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Non"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Oui"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Qté :"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Champignons"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   450
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Non"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Oui"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Qté :"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "Manger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.QtéChampignon = Form1.QtéChampignon - Val(List1.List(List1.ListIndex))
    Form1.QtéFruit = Form1.QtéFruit - Val(List2.List(List2.ListIndex))
    Form1.QtéPoissons = Form1.QtéPoissons - Val(List3.List(List3.ListIndex))
    Life = Life + (Val(List1.List(List1.ListIndex)) * 6) 'Ajout du nombre de Champi
    Life = Life + (Val(List2.List(List2.ListIndex)) * 4) 'Ajout du nombre de Fruis
    Life = Life + (Val(List3.List(List3.ListIndex)) * 9) 'Ajout du nombre de Poissons
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    List1.Clear
    List2.Clear
    List3.Clear
    For i = 0 To Form1.QtéChampignon
        List1.AddItem i
    Next i
    For i = 0 To Form1.QtéFruit
        List2.AddItem i
    Next i
    For i = 0 To Form1.QtéPoissons
        List3.AddItem i
    Next i
End Sub
