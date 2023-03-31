VERSION 5.00
Begin VB.Form FrmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox TxtTemps 
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "10"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Regarder si les dés sont lancés toute les"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "secondes"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    If Int(TxtTemps.Text) < 5 Then
        MsgBox "La valeur doit être plus grande que 4", vbCritical, "Erreur"
    Else
        FrmMain.Show
        FrmMain.ClkWebcam.Enabled = True
        FrmOption.Hide
    End If
End Sub
