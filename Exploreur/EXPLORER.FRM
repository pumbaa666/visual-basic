VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorateur"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CmdOuvrir 
      Caption         =   "&Ouvrir"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtChemin 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Chemin"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    End
End Sub

Private Sub CmdOuvrir_Click()
Dim var As Double
    If Dir1.Path = "c:\" Or Dir1.Path = "C:\" Then
        var = Shell("C:\WINDOWS\NOTEPAD.EXE " & Dir1.Path & File1.FileName, vbNormalFocus)
    Else
        var = Shell("C:\WINDOWS\NOTEPAD.EXE " & Dir1.Path & "\" & File1.FileName, vbNormalFocus)
    End If
End Sub

Private Sub Dir1_Change()
    TxtChemin.Text = Dir1.Path
    File1.Path = Dir1.Path
End Sub

Private Sub File1_DblClick()
    CmdOuvrir_Click
End Sub

Private Sub File1_Click()
    If Dir1.Path = "c:\" Or Dir1.Path = "C:\" Then
        TxtChemin.Text = TxtChemin.Text & File1.FileName
    Else
        TxtChemin.Text = Dir1.Path & "\" & File1.FileName
    End If
End Sub

Private Sub Form_Load()
    Dir1.Path = "c:\temp"
End Sub

Private Sub TxtChemin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dir1.Path = TxtChemin.Text
    End If
End Sub
