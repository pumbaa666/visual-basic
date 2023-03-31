VERSION 5.00
Begin VB.Form FrmListe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   2220
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Liste 
      Height          =   4545
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Temps"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "FrmListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Liste_Click(Index As Integer)
Dim vCount As Integer

    For vCount = 0 To vNbEntree
        Liste(vCount).ListIndex = Liste(Index).ListIndex
    Next
End Sub
