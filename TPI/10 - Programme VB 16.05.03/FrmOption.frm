VERSION 5.00
Begin VB.Form FrmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTemps 
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "10"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox TxtNbDes 
      Height          =   285
      Left            =   3120
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "4"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label LblSec 
      Caption         =   "seconde(s)  [1 --> 99]"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Regarder si les dés sont lancés toute les"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre de dés à lancer en un jet : "
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Score maximum à atteindre pour gagner"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "[1 --> 4]"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label LblPlage 
      Caption         =   "[1 --> 4]"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    If TxtTemps.Text = "" Then
        MsgBox "Le temps entre chaque image prise par la webcam doit être plus d'au moins une seconde", vbCritical, "Erreur"
    ElseIf Int(TxtTemps.Text) = 0 Then
        MsgBox "Le temps entre chaque image prise par la webcam doit être plus d'au moins une seconde", vbCritical, "Erreur"
    ElseIf TxtNbDes.Text = "" Then
        MsgBox "Il faut lancer au moins un dé", vbCritical, "Erreur"
    ElseIf Int(TxtNbDes.Text) > 4 Then
        MsgBox "3 dés maximum", vbCritical, "Erreur"
    ElseIf Int(TxtNbDes.Text) = 0 Then
        MsgBox "Il faut lancer au moins un dé", vbCritical, "Erreur"
    ElseIf TxtScore.Text = "" Then
        MsgBox "Le score à atteindre doit être suppérieur à zéro", vbCritical, "Erreur"
    ElseIf Int(TxtScore.Text) > (Int(TxtNbDes.Text) * 6) * 2 / 3 Then
        MsgBox "Le score à atteindre doit être inférieur ou égal aux 2/3 de la valeur maxiumum atteignable par tout les dés (ici :" & (Int(TxtNbDes.Text) * 6) * 2 / 3 & ")", vbCritical, "Erreur"
    ElseIf Int(TxtScore.Text) = 0 Then
        MsgBox "Le score à atteindre doit être suppérieur à 0", vbCritical, "Erreur"
    Else
        FrmMain.Show
        FrmMain.ClkWebcam.Enabled = True
        FrmOption.Hide
    End If
End Sub

Function ProtectionSaisie(ByVal vKeyAscii)
    If (vKeyAscii < 48 Or vKeyAscii > 57) And vKeyAscii <> 8 And vKeyAscii <> 13 Then
        MsgBox "Veuillez n'entrer que des chiffres", vbCritical, "Erreur"
        ProtectionSaisie = 0
    Else
        If vKeyAscii = 13 Then
            CmdOk_Click
        End If
        ProtectionSaisie = vKeyAscii
    End If
End Function

Private Sub TxtNbDes_KeyPress(KeyAscii As Integer)
    KeyAscii = ProtectionSaisie(KeyAscii)
End Sub

Private Sub TxtNbDes_LostFocus()
    If TxtNbDes.Text <> "" Then
        LblPlage.Caption = "[1 --> " & (Int(TxtNbDes.Text) * 6) * 2 / 3 & "]"
    End If
End Sub

Private Sub TxtScore_KeyPress(KeyAscii As Integer)
    KeyAscii = ProtectionSaisie(KeyAscii)
End Sub

Private Sub TxtTemps_KeyPress(KeyAscii As Integer)
    KeyAscii = ProtectionSaisie(KeyAscii)
End Sub
