VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   1080
   ClientTop       =   390
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Calculer"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuitter_Click()
    End
End Sub

Private Sub cmdOk_Click()

    If txt(0).Text = "" Then
        If txt(2).Text <> "" Then
            If txt(1).Text <> "" Then
                txt(0).Text = (100 * Val(txt(2).Text)) / (100 + Val(txt(1).Text))
            Else
                MsgBox "Entrez la valeure initiale ou le pourcentage", vbCritical, "Erreur"
            End If
        ElseIf txt(1).Text <> "" Then
            MsgBox "Entrez la valeure initiale ou le total", vbCritical, "Erreur"
        Else
            MsgBox "Entrez au moins 2 valeurs", vbCritical, "Erreur"
        End If
    ElseIf txt(2).Text = "" Then
        If txt(1).Text <> "" Then
            txt(2).Text = Val(txt(0).Text) + Val(txt(0).Text) * Val(txt(1).Text) / 100
        Else
            MsgBox "Entrez le pourcentage ou le total", vbCritical, "Erreur"
        End If
    Else
        If txt(1) = "" Then
            If Val(txt(0).Text) <> 0 Then
                txt(1) = (Val(txt(2).Text) - Val(txt(0).Text)) / Val(txt(0).Text) * 100
            Else
                txt(1).Text = "Div 0"
            End If
        Else
            MsgBox "N'entrez que 2 valeurs", vbCritical, "Erreur"
        End If
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOk_Click
    End If
End Sub
