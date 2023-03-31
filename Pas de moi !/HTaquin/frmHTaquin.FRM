VERSION 5.00
Begin VB.Form Taquin 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFEECC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTaquin : Jeu de Taquin"
   ClientHeight    =   3465
   ClientLeft      =   1320
   ClientTop       =   1695
   ClientWidth     =   4905
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmHTaquin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox btnScores 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3450
      Picture         =   "frmHTaquin.frx":030A
      ScaleHeight     =   360
      ScaleWidth      =   1260
      TabIndex        =   8
      Top             =   75
      Width           =   1260
   End
   Begin VB.PictureBox CFin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3450
      Picture         =   "frmHTaquin.frx":099F
      ScaleHeight     =   360
      ScaleWidth      =   1260
      TabIndex        =   7
      Top             =   2850
      Width           =   1260
   End
   Begin VB.PictureBox CNouv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3450
      Picture         =   "frmHTaquin.frx":104D
      ScaleHeight     =   360
      ScaleWidth      =   1260
      TabIndex        =   6
      Top             =   2400
      Width           =   1260
   End
   Begin VB.CommandButton Piece 
      BackColor       =   &H00FFEECC&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   240
      MaskColor       =   &H00FFEECC&
      TabIndex        =   5
      Top             =   240
      Width           =   750
   End
   Begin VB.HScrollBar AHori 
      Height          =   255
      Left            =   3480
      Max             =   8
      Min             =   3
      TabIndex        =   0
      Top             =   2040
      Value           =   4
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   0
   End
   Begin VB.Label Ifond 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                                      Mon Taquin"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   225
      TabIndex        =   2
      Top             =   225
      Width           =   3030
      WordWrap        =   -1  'True
   End
   Begin VB.Label NPieces 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3915
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Duree 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEECC&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pièces par côté"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Taquin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, t, x, y, x1, y1 As Integer
Dim iPos As Integer
Dim sH, sH1, sName, sScore As String
Dim iMax, iSize             ' Nombre de pièces, largeur/hauteur
Dim lHeureDemarr As Date            ' Heure de démarrage
Dim tiChamp()                ' Champ de stockage des positions
Private Sub AHori_Change()
    NPieces.Caption = Str(AHori.Value)
End Sub
Private Sub CFin_Click()
    Unload Scores
    ' Sauver les scores
    Open App.Path & "\HTaquin.sco" For Output As #1
    For t = 1 To 6
        For i = 1 To 10
            Print #1, tsHigh(t, i)
        Next i
    Next t
    Close #1
    End
End Sub
Private Sub CNouv_Click()
    Taquin.MousePointer = 11
    ' Effacer les pièces en cours
    For i = 1 To iMax * iMax - 1
        Unload Piece(i)
    Next
    
    iMax = AHori.Value
    Dimensionner
    
    ' Distribuer les pièces dans le champ
    For y = 0 To iMax - 1
        For x = 0 To iMax - 1
            i = x + y * iMax
            If i < iMax * iMax - 1 Then
                Piece(i).Width = iSize
                Piece(i).Height = iSize
                Piece(i).Top = (y * iSize) + Ifond.Top
                Piece(i).Left = (x * iSize) + Ifond.Left
                Piece(i).Visible = 1
                Piece(i).Caption = ""
            End If
        Next x
    Next y

    Taquin.Show

'   Numéroter
    tiChamp(iMax, iMax) = 0
    For i = 1 To iMax * iMax - 1
        Do
            j = Rnd * (iMax * iMax - 2)
        Loop Until Piece(j).Caption = ""
        Piece(j).Caption = Trim(Str(i))
        tiChamp(j Mod iMax + 1, j \ iMax + 1) = i
    Next

    Taquin.MousePointer = 0
    lHeureDemarr = Now
    Timer1.Enabled = True

End Sub
Private Sub Dimensionner()
    ' Dimensionner le champ
    iSize = 3000 / iMax
    ReDim tiChamp(iMax + 1, iMax + 1)
    Select Case AHori.Value
        Case 3: Piece(0).FontSize = 22
        Case 4: Piece(0).FontSize = 20
        Case 5: Piece(0).FontSize = 18
        Case 6: Piece(0).FontSize = 14
        Case 7: Piece(0).FontSize = 11
        Case 8: Piece(0).FontSize = 10
    End Select
    Piece(0).Width = iSize
    Piece(0).Height = iSize
    Piece(0).Top = Ifond.Top
    Piece(0).Left = Ifond.Left + 30
    Piece(0).Visible = 0
    
    ' Propager la 1ère pièce
    For i = 1 To iMax * iMax - 1
        Load Piece(i)
    Next
End Sub
Private Sub Form_Load()
    Randomize Timer

    ' Afficher le nombre de pièces initial
    NPieces.Caption = Str(AHori.Value)

    iMax = 8
    Dimensionner

    ' Lecture / Création du fichier des scores
    If Len(Dir$(App.Path & "\HTaquin.sco")) Then
        Open App.Path & "\HTaquin.sco" For Input As #1
        For t = 1 To 6
            For i = 1 To 10
                Line Input #1, tsHigh(t, i)
            Next i
        Next t
        Close #1
    Else
        Open App.Path & "\HTaquin.sco" For Output As #1
        For t = 1 To 6
            For i = 1 To 10
                tsHigh(t, i) = "oo:oo:oo  oo/oo/oo  Alfred NONYME"
                Print #1, tsHigh(t, i)
            Next i
        Next t
        Close #1
    End If
    
End Sub
Private Sub btnScores_Click()
    Scores.Show
End Sub
Private Sub Piece_Click(Index As Integer)
    ' Déplacement éventuel d'une pièce
    x1 = (Piece(Index).Left - Ifond.Left) \ iSize
    y1 = (Piece(Index).Top - Ifond.Top) \ iSize

    ' Renseigner la position dans le champ
    x = x1 + 1: y = y1 + 1
    If x > 1 And tiChamp(x - 1, y) = 0 Then x1 = x1 - 1
    If x < iMax And tiChamp(x + 1, y) = 0 Then x1 = x1 + 1
    If y > 1 And tiChamp(x, y - 1) = 0 Then y1 = y1 - 1
    If y < iMax And tiChamp(x, y + 1) = 0 Then y1 = y1 + 1

    ' Déplacement éventuel
    If x = x1 + 1 And y = y1 + 1 Then Exit Sub
    Piece(Index).Move (x1 * iSize) + Ifond.Left, (y1 * iSize) + Ifond.Top
    tiChamp(x, y) = 0
    tiChamp(x1 + 1, y1 + 1) = Val(Piece(Index).Caption)

    ' Test du champ entier pour détermination du gagné
    For i = 0 To iMax * iMax - 2
        If tiChamp(i Mod iMax + 1, i \ iMax + 1) <> i + 1 Then Exit For
    Next
    If i = iMax * iMax - 1 Then
        Beep
        Timer1.Enabled = 0
        MsgBox "C'est gagné", 48

        ' Cherche si meilleur score
        iPos = 0
        sH1 = Duree.Caption
        For i = 1 To 10
            sH = tsHigh(AHori.Value - 2, i)
            If sH > sH1 Then iPos = i: sH1 = sH
        Next i
        If iPos Then
            sName = InputBox( _
                    "Votre nom SVP", _
                    "Un très bon score ...", _
                    "Alfred NONYME")
            sScore = Duree.Caption + "  " & _
                    Mid(Date, 4, 2) + "/" & _
                    Left(Date$, 2) + "/" & _
                    Right$(Date$, 2) + "  " & sName
            tsHigh(AHori.Value - 2, iPos) = sScore
            If Scores.Visible Then Scores.NbrePieces.SetFocus
        Else
            MsgBox _
                "Vous n'êtes pas au tableau d'honneur cette fois-ci ! Dommage !"
        End If

        CNouv.SetFocus
    End If
End Sub
Private Sub Timer1_Timer()
    Duree.Caption = Format(Now - lHeureDemarr, "hh:mm:ss")
End Sub

