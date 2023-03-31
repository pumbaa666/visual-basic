VERSION 5.00
Begin VB.Form FrmListe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mes DVD's - Liste"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtAffichage 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   13095
   End
   Begin VB.ListBox Liste 
      Height          =   4935
      Index           =   5
      Left            =   10800
      TabIndex        =   5
      ToolTipText     =   "Prêté"
      Top             =   840
      Width           =   2535
   End
   Begin VB.ListBox Liste 
      Height          =   4935
      Index           =   4
      Left            =   9960
      TabIndex        =   4
      ToolTipText     =   "Note"
      Top             =   840
      Width           =   735
   End
   Begin VB.ListBox Liste 
      Height          =   4935
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Titre"
      Top             =   840
      Width           =   3255
   End
   Begin VB.ListBox Liste 
      Height          =   4935
      Index           =   0
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Numéro"
      Top             =   840
      Width           =   735
   End
   Begin VB.ListBox Liste 
      Height          =   4935
      Index           =   2
      Left            =   4440
      TabIndex        =   1
      ToolTipText     =   "Genre"
      Top             =   840
      Width           =   2415
   End
   Begin VB.ListBox Liste 
      Height          =   4935
      Index           =   3
      Left            =   6960
      TabIndex        =   0
      ToolTipText     =   "Acteurs"
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Prêté à"
      Height          =   255
      Left            =   10800
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Note"
      Height          =   255
      Left            =   9960
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Acteur(s)"
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Genre"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Titre"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Numéro"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Visible         =   0   'False
      Begin VB.Menu FichierPreter 
         Caption         =   "Prêter"
      End
      Begin VB.Menu FichierReprendre 
         Caption         =   "Reprendre"
      End
      Begin VB.Menu FichierTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu FichierModifier 
         Caption         =   "Modifier"
      End
      Begin VB.Menu FichierSupprimer 
         Caption         =   "Supprimer"
      End
   End
End
Attribute VB_Name = "FrmListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FichierPreter_Click()
    BouttonPreter
    FrmPreter.TxtNum.Text = Liste(0).Text
    FrmPreter.TxtNom.SetFocus
End Sub

Private Sub FichierReprendre_Click()
Dim vTemp As Integer

    tListe(5, Int(Liste(0).Text)) = ""
    Preter Int(Liste(0).Text), ""
    vTemp = Liste(5).ListIndex
    Liste(5).RemoveItem vTemp
    Liste(5).AddItem "", vTemp
End Sub

Private Sub FichierSupprimer_Click()
    DelDVD Liste(0).Text
End Sub

Private Sub FichierModifier_Click()
    BouttonModifier
'    Liste(0).ListIndex = Int(Liste(0).Text) - 1
    FrmAdd.TxtNum.Text = Int(Liste(0).Text)
End Sub

Private Sub Liste_Click(Index As Integer)
    If vTestListe5 = True Then
        Liste_MouseUp Index, 1, 0, 0, 0
    End If
End Sub

Private Sub Liste_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Liste_MouseUp Index, 1, 0, 0, 0
End Sub

Private Sub Liste_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Liste_MouseUp Index, Button, Shift, X, Y
End Sub

Private Sub Liste_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vCount As Integer
Dim vTo As Integer
'Static vTest As Boolean

'    If vTest = False Then
'        vTest = True
        vTestListe5 = False
        TxtAffichage.Text = Liste(Index).Text
        If Button = 1 Then
            If Liste(5).ListCount <> 0 Then
                vTo = 5
            Else
                vTo = 4
            End If
            For vCount = 0 To vTo
                Liste(vCount).ListIndex = Liste(Index).ListIndex
            Next
        Else
            If Liste(0).ListIndex = -1 Then
                MsgBox "Veuillez sélectionner un DVD", vbCritical, "Erreur"
            Else
                If Liste(0).ListCount <> vNbDVDTot Then
                    MsgBox "Veuillez afficher toute la liste", vbCritical, "Erreur"
                Else
                    If Liste(5).Text = "" Then
                        FichierReprendre.Enabled = False
                        FichierPreter.Enabled = True
                    Else
                        FichierReprendre.Enabled = True
                        FichierPreter.Enabled = False
                    End If
                    PopupMenu Fichier
                End If
            End If
        End If
        vTestListe5 = True
'    End If
'    vTest = False
End Sub
