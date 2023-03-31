VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajouter"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Retour"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox TxtEmail 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox TxtDate 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "JJ.MM.AAAA"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TxtPrenom 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TxtNom 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "E-mail (facultatif)"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Date de naissance"
      Height          =   255
      Left            =   -120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Prénom"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nom"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
Dim vDate As Date
Dim vTemp As String

    On Error GoTo ErrorDate
    If TxtNom.Text = "" Then
        MsgBox "Veuillez entrer le nom", vbCritical
    ElseIf TxtPrenom.Text = "" Then
        MsgBox "Veuillez entrer le prénom", vbCritical
    ElseIf TxtDate.Text = "" Then
        MsgBox "Veuillez entrer la date", vbCritical
    ElseIf Len(TxtDate.Text) <> 10 Then
        MsgBox "Le format de la date n'est pas correct!", vbCritical
        vDate = "01.01.1980"
    Else
        vDate = TxtDate.Text
        
        Open "c:\temp\donnees.dat" For Input As #1
        Line Input #1, vTemp
        Close #1
        
        If vTemp = "          " Then
            Open "c:\temp\donnees.dat" For Output As #1
        Else
            Open "c:\temp\donnees.dat" For Append As #1
        End If
        Print #1, TxtNom.Text
        Print #1, TxtPrenom.Text
        Print #1, TxtDate.Text
        Print #1, TxtEmail.Text
        Close #1
        MsgBox "Ajout correctement effectué", vbInformation
    End If
ErrorDate:
    If vDate = 0 Then
        MsgBox "Le format de la date n'est pas correct!", vbCritical
    End If
End Sub

Private Sub CmdQuitter_Click()
    FrmAdd.Hide
    FrmMain.Show
End Sub

Private Sub TxtDate_KeyPress(KeyAscii As Integer)
    TxtNom_KeyPress (KeyAscii)
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    TxtNom_KeyPress (KeyAscii)
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        MsgBox "Non non, on évite les espaces !!!", vbCritical
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtPrenom_KeyPress(KeyAscii As Integer)
    TxtNom_KeyPress (KeyAscii)
End Sub
