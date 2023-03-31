VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MadMatt"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkTest 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5040
      Top             =   3960
   End
   Begin VB.CommandButton CmdListe 
      Caption         =   "Liste"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox NumberStart 
      Height          =   615
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Arrêter la recherche"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Start 
      Caption         =   "Démarrer la recherche des nombres premiers"
      Default         =   -1  'True
      Height          =   735
      Left            =   6480
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Texte 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.ListBox Liste 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Début de la recherche :"
      Height          =   195
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StopSearch As Boolean
Dim NombreEnCours As Double

Private Sub ClkTest_Timer()
Static vCount As Integer
    If vCount = 0 Then
        vCount = 1
    Else
        ClkTest.Enabled = False
        StopSearch = True
    End If
End Sub

Private Sub CmdListe_Click()
    If Liste.Visible = True Then
        Liste.Visible = False
    Else
        Liste.Visible = True
    End If
End Sub

' On charge le dernier nombre premier trouvé
Private Sub Form_Load()
    ChDir App.Path
    NombreEnCours = 1#
    Load
    NumberStart.Text = Trim(Str(NombreEnCours))
End Sub

' Affiche le nombre premier quand on clique dans la liste
Private Sub Liste_Click()
    Texte.Text = Trim(Liste.List(Liste.ListIndex))
End Sub

' Démarre la recherche
Private Sub Start_Click()
'    ClkTest.Enabled = True
    CherchePremier
End Sub

' Procédure qui cherche les nombres premiers
Private Function CherchePremier()
    Dim Number As Double
    Me.Caption = "Recherche ... - MadMatt"
    ' Vide la liste
    Liste.Clear
    ' Initialise tout
    StopSearch = False
    btnStop.Enabled = True
    Start.Enabled = False
Debut:
    Number = NombreEnCours
    ' cherche
    Do
        If StopSearch = True Then GoTo Fin
        ' Afin de ne pas stocker trop de nombres dans la listbox on quitte et on revient
        If Number - NombreEnCours > 99999 Then
            NombreEnCours = Number
            NumberStart.Text = Trim(Str(NombreEnCours))
            Save
            Liste.Clear
            GoTo Debut
        End If
        If Premier(Number) Then Liste.AddItem Str(Number)
        Number = Number + 1#
        DoEvents
    Loop
Fin:
    btnStop.Enabled = False
    Start.Enabled = True
    NombreEnCours = Number
    Save
    NumberStart.Text = Trim(Str(NombreEnCours))
    Me.Caption = "Fini !"
End Function

' Renvoie True si le nombre passé en parametre est premier
Private Function Premier(ByVal Number As Double) As Boolean
    Dim T As Double, A As Double
    ' Un nombre premier régit forcément à ces regles
    If ((Number Mod 4 = 3 Or Number Mod 4 = 1) And (Number Mod 6 = 1 Or Number Mod 6 = 5)) Then
        ' On teste la primalité - la nombre est peut etre premier
        ' On divise le nombre par tous les nombres inférieurs à lui jusqu'a sa racine
        ' (pas besoin d'aller plus loin merci Pingouin)
        For T = 2# To Int(Sqr(Number)) + 1#
            A = Number / T
            ' Si la division est entière, alors le nombre a un diviseur (a)
            If A = Int(A) Then
                Premier = False
                Exit Function
            End If
            'DoEvents
        Next
    Else
        Premier = False
        Exit Function
    End If
    Premier = True
End Function

' Quand on quitte on demande si on veut sauvegarder
Private Sub Form_Unload(Cancel As Integer)
    If btnStop.Enabled = True Then
        A = MsgBox("Voulez-vous sauvegarder vos derniers résultats ?", vbQuestion + vbYesNo, "Attention")
        If A = vbYes Then Save
    End If
    End
End Sub

Private Sub btnStop_Click()
    StopSearch = True
End Sub

' Sauvegarde tous les entiers trouvés et le dernier nombre traité
Private Function Save()
    'On Error GoTo Fin
    Dim T As Double
    Open "Nombres.txt" For Append As #1
    For T = 0# To Liste.ListCount - 1#
        Print #1, Trim(Liste.List(T))
    Next
    Close #1
    Open "En cours.txt" For Output As #1
    Print #1, Trim(Str(NombreEnCours))
Fin:
    Close #1
End Function

' Charge le dernier nombre traité
Private Function Load()
    On Error GoTo Fin
    Dim T As Double
    Open "En cours.txt" For Input As #1
    NombreEnCours = CDbl(Input(LOF(1), #1))
    Close #1
Fin:
End Function
