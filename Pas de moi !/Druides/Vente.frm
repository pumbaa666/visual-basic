VERSION 5.00
Begin VB.Form Vente 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2670
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Fruits"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Acheter"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vendre"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Taux pour 1 : "
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Qté :"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Champignons"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Acheter"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Vendre"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Taux pour 1 : "
         Height          =   195
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Qté :"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Valider"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
End
Attribute VB_Name = "Vente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option2(0).Value = True Then 'Champignons
        If Form1.QtéChampignon < Val(Text2.Text) Then
            MsgBox "Vous n'avez pas assez de champignon à vendre", vbExclamation, "Le Voyageur"
        Else
            Form1.Somme = Form1.Somme + (Val(Text2.Text) * Val(Label4.Caption))
            Form1.QtéChampignon = Form1.QtéChampignon - Val(Text2.Text)
        End If
    Else
        If (Val(Text1.Text) * Val(Label6.Caption)) > Form1.Somme Then
            MsgBox "Vous n'avez pas assez d'argent pour payer les champignons", vbExclamation, "Le voyageur"
        Else
            Form1.Somme = Form1.Somme - (Val(Text1.Text) * Val(Label6.Caption))
            Form1.QtéChampignon = Form1.QtéChampignon + Val(Text1.Text)
        End If
    End If

    If Option1(1).Value = True Then 'Fruits
        If Form1.QtéFruit < Val(Text2.Text) Then
            MsgBox "Vous n'avez pas assez de fruits à vendre", vbExclamation, "Le Voyageur"
        Else
            Form1.Somme = Form1.Somme + (Val(Text2.Text) * Val(Label6.Caption))
            Form1.QtéFruit = Form1.QtéFruit - Val(Text2.Text)
        End If
    Else
        If (Val(Text2.Text) * Val(Label4.Caption)) > Form1.Somme Then
            MsgBox "Vous n'avez pas assez d'argent pour payer les fruits", vbExclamation, "Le voyageur"
        Else
            Form1.Somme = Form1.Somme - (Val(Text2.Text) * Val(Label4.Caption))
            Form1.QtéFruit = Form1.QtéFruit + Val(Text2.Text)
        End If
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Label4.Caption = Int((Rnd * 2) + 1)
    Else
        Label4.Caption = Int((Rnd * 4) + 2)
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    If Index = 0 Then
        Label6.Caption = Int((Rnd * 3) + 1)
    Else
        Label6.Caption = Int((Rnd * 6) + 3)
    End If
End Sub
