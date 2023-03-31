VERSION 5.00
Begin VB.Form FrmChanger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changer d'emplacement"
   ClientHeight    =   2505
   ClientLeft      =   1545
   ClientTop       =   675
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox ComboCouleur 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox TxtNom 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox ComboPlace 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "ComboPlace"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label LblTxt 
      Alignment       =   2  'Center
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image ImageVéhicule 
      Height          =   1935
      Left            =   4320
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Emplacement : "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Couleur : "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nom : "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnnuler_Click()
    FrmChanger.Hide
End Sub

Private Sub CmdOk_Click()
Dim i As Integer
Dim vString As String
Dim vCouleur As String

    If Len(TxtNom.Text) <> 9 Then
        MsgBox "Le nom n'est pas correct. Il doit être de la forme : XXXX 9999", vbCritical, "Erreur"
    Else
        Open CheminRelatif & "\pss 90.txt" For Input As #1
        Open CheminRelatif & "\Copie de pss 90.txt" For Output As #2
        i = 0
        Do
            Line Input #1, vString
            If i <> tEnCours(3) Then
                Print #2, vString
            Else
                vCouleur = ComboCouleur.Text
                If Len(ComboCouleur.Text) < 5 Then
                    vCouleur = ComboCouleur.Text & " "
                End If
                Print #2, TxtNom.Text & "," & vCouleur & "," & ComboPlace.Text
            End If
            i = i + 1
        Loop Until (EOF(1))
        Close 2
        Close 1
        Kill CheminRelatif & "\pss 90.txt"
        Name CheminRelatif & "\Copie de pss 90.txt" As CheminRelatif & "\pss 90.txt"
        FrmChanger.Hide
    End If
End Sub

Private Sub ComboPlace_Click()
    If IsNumeric(ComboPlace.Text) Then
        ImageVéhicule.Picture = LoadPicture(CheminRelatif & "images\" & ComboPlace.Text & ".jpg")
        ImageVéhicule.Visible = True
        LblTxt.Visible = False
    Else
        LblTxt.Caption = ComboPlace.Text
        LblTxt.Visible = True
        ImageVéhicule.Visible = False
    End If
End Sub

Private Sub Form_Activate()
Dim i As Integer

    TxtNom.Text = tEnCours(0)

    ComboCouleur.Clear
    ComboCouleur.AddItem "Rouge"
    ComboCouleur.AddItem "Vert"
    ComboCouleur.AddItem "Bleu"
    ComboCouleur.Text = tEnCours(1)

    
    ComboPlace.Clear

    ComboPlace.AddItem "Local/Réserve"
    ComboPlace.AddItem "Révision"
    ComboPlace.AddItem "Défectueux"
    ComboPlace.AddItem "En prêt"
    ComboPlace.AddItem "Cours/Exercice"
    ComboPlace.Text = tEnCours(2)
    ComboPlace_Click

    ComboPlace.AddItem "005"
    ComboPlace.AddItem "300"
    ComboPlace.AddItem "301"
    ComboPlace.AddItem "305"
    ComboPlace.AddItem "306"
    ComboPlace.AddItem "312"
    ComboPlace.AddItem "315"
    ComboPlace.AddItem "320"
    ComboPlace.AddItem "331"
    ComboPlace.AddItem "360"
    ComboPlace.AddItem ""
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub
