VERSION 5.00
Begin VB.Form FrmDel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supprimer"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "Sélectionnez la personne à supprimer..."
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Retour"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   4695
   End
End
Attribute VB_Name = "FrmDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuitter_Click()
    FrmDel.Hide
    FrmMain.Show
End Sub

Private Sub Combo_Click()
    fSupprimer (Combo.ListIndex)
End Sub

Private Sub Form_Activate()
Dim vData(4) As String
Dim vChaineTot As String
Dim vCount As Integer

    Combo.Text = "Sélectionnez la personne à supprimer..."
    Open "c:\temp\donnees.dat" For Input As #1
    Do
        Line Input #1, vData(vCount)
        vChaineTot = vChaineTot & vData(vCount) & "    "
        vCount = vCount + 1
        If vCount = 4 Then
            vCount = 0
            Combo.AddItem vChaineTot
            vChaineTot = ""
        End If
    Loop Until (EOF(1))
    Close #1
End Sub
