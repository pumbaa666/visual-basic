VERSION 5.00
Begin VB.Form FrmTable 
   Caption         =   "Table Ascii"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   7
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   6
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   5
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox LstTable 
      Height          =   6495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim vCount As Long
    For vCount = 0 To 255
        LstTable(Int(vCount / 32)).AddItem vCount & ": " & Chr(vCount)
    Next
End Sub


