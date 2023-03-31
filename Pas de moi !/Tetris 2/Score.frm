VERSION 5.00
Begin VB.Form meilleurs_scores 
   BackColor       =   &H00FF8080&
   Caption         =   "Meilleurs Scores du Brick Game"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "Score.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton fermage 
      BackColor       =   &H00FF8080&
      Caption         =   "Fermer"
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   35
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   34
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   33
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   32
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   31
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   30
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   29
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Position"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   1080
      Y1              =   240
      Y2              =   5520
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   24
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   22
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   21
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   17
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   16
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   15
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Score 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Nom 
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Score"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Nom du joueur"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   240
      Y2              =   5520
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "meilleurs_scores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fermage_Click()

Unload Maine
Load info_partie
info_partie.Show
Unload meilleurs_scores
End Sub

Private Sub Form_Load()


    'Déclaration des variables
    Dim i As Integer
    Dim j As Integer
    Dim Score_partie As Integer
    Dim score_dans_tableau As t_score
    Score_partie = Maine.Score_val.Caption
    'Introduit dans le tableau des scores, le score obtenu
    tablo(10).Score = Score_partie
    tablo(10).Nom = info_partie.Nom_joueur.Text


    'Répéte pour chaque cases du tableau
    For i = 9 To 0 Step -1

        'Ne prend pas en compte le plus grand score
        For j = 9 To 9 - i Step -1

            'Inverse les scores si le score en-dessous est plus grand que celui en dessus
            If tablo(j + 1).Score >= tablo(j).Score Then

                score_dans_tableau = tablo(j)
                tablo(j) = tablo(j + 1)
                tablo(j + 1) = score_dans_tableau
                
   

    Nom(0).Caption = tablo(0).Nom
    Score(0).Caption = tablo(0).Score
    Nom(1).Caption = tablo(1).Nom
    Score(1).Caption = tablo(1).Score
    Nom(2).Caption = tablo(2).Nom
    Score(2).Caption = tablo(2).Score
    Nom(3).Caption = tablo(3).Nom
    Score(3).Caption = tablo(3).Score
    Nom(4).Caption = tablo(4).Nom
    Score(4).Caption = tablo(4).Score
    Nom(5).Caption = tablo(5).Nom
    Score(5).Caption = tablo(5).Score
    Nom(6).Caption = tablo(6).Nom
    Score(6).Caption = tablo(6).Score
    Nom(7).Caption = tablo(7).Nom
    Score(7).Caption = tablo(7).Score
    Nom(8).Caption = tablo(8).Nom
    Score(8).Caption = tablo(8).Score
    Nom(9).Caption = tablo(9).Nom
    Score(9).Caption = tablo(9).Score
    Nom(10).Caption = tablo(10).Nom
    Score(10).Caption = tablo(10).Score
    meilleurs_scores.Show
             End If
        Next j
    Next i
    Call Score_inscrir
    

End Sub

Public Sub Score_inscrir()

    On Error GoTo titi
    
    Open "Jacqueline.txt" For Output As #1
    Write #1, tablo(0).Nom
    Write #1, tablo(0).Score
    Write #1, tablo(1).Nom
    Write #1, tablo(1).Score
    Write #1, tablo(2).Nom
    Write #1, tablo(2).Score
    Write #1, tablo(3).Nom
    Write #1, tablo(3).Score
    Write #1, tablo(4).Nom
    Write #1, tablo(4).Score
    Write #1, tablo(5).Nom
    Write #1, tablo(5).Score
    Write #1, tablo(6).Nom
    Write #1, tablo(6).Score
    Write #1, tablo(7).Nom
    Write #1, tablo(7).Score
    Write #1, tablo(8).Nom
    Write #1, tablo(8).Score
    Write #1, tablo(9).Nom
    Write #1, tablo(9).Score
 '   Write #1, Nom(10).Caption
'    Write #1, Score(10).Caption

    Close #1
titi:
Close #1
End Sub
Public Sub Score_lire()

    On Error GoTo titi
    
    Open "Jacqueline.txt" For Input As #1
    Input #1, tablo(0).Nom
    Input #1, tablo(0).Score
    Input #1, tablo(1).Nom
    Input #1, tablo(1).Score
    Input #1, tablo(2).Nom
    Input #1, tablo(2).Score
    Input #1, tablo(3).Nom
    Input #1, tablo(3).Score
    Input #1, tablo(4).Nom
    Input #1, tablo(4).Score
    Input #1, tablo(5).Nom
    Input #1, tablo(5).Score
    Input #1, tablo(6).Nom
    Input #1, tablo(6).Score
    Input #1, tablo(7).Nom
    Input #1, tablo(7).Score
    Input #1, tablo(8).Nom
    Input #1, tablo(8).Score
    Input #1, tablo(9).Nom
    Input #1, tablo(9).Score
'    Input #1, tablo(10).Nom
'    Input #1, tablo(10).Score
    Close #1
titi:
Close #1
End Sub
