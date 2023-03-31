VERSION 5.00
Begin VB.Form FrmShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulter"
   ClientHeight    =   3210
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   1620
      Index           =   3
      Left            =   4800
      TabIndex        =   8
      Top             =   720
      Width           =   2655
   End
   Begin VB.ListBox List 
      Height          =   1620
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List 
      Height          =   1620
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdTri 
      Caption         =   "&Adresse e-mail"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton CmdTri 
      Caption         =   "&Date"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton CmdTri 
      Caption         =   "&Prénom"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton CmdTri 
      Caption         =   "&Nom"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Retour"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   7095
   End
   Begin VB.ListBox List 
      Height          =   1620
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Visible         =   0   'False
      Begin VB.Menu MenuFichierDel 
         Caption         =   "Supprimer"
      End
      Begin VB.Menu MenuFichierMail 
         Caption         =   "Envoyer un mail"
      End
   End
End
Attribute VB_Name = "FrmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tDonnee(4, 50) As String

Private Sub CmdQuitter_Click()
    FrmShow.Hide
    FrmMain.Show
End Sub

Private Sub CmdTri_Click(Index As Integer)
Dim vCount As Integer
Dim vCount2 As Integer
Dim vCount3 As Integer
Dim vCount4 As Integer
Dim vTemp2 As String
Dim tTemp(4) As String

On Error Resume Next
    For vCount = 0 To List(0).ListCount - 1
        For vCount2 = vCount To List(0).ListCount - 1
            If vCount2 = vCount Then
                For vCount3 = 0 To 4
                    tTemp(vCount3) = tDonnee(vCount3, vCount2)
                Next
            Else

        ' Tri de la date selon la méthode choisie dans le menu
                If Index = 2 Then
                    If CDate(tDonnee(2, vCount2)) < CDate(tTemp(Index)) Then
                        For vCount3 = 0 To 4
                            tTemp(vCount3) = tDonnee(vCount3, vCount2)
                        Next
                    End If
                Else
                    If tDonnee(Index, vCount2) < tTemp(Index) Then
                        For vCount3 = 0 To 4
                            tTemp(vCount3) = tDonnee(vCount3, vCount2)
                        Next
                    End If
                End If
            End If
        Next

' Fais l'inversion
        For vCount4 = 0 To 3
            vTemp2 = tDonnee(vCount4, vCount)
            tDonnee(vCount4, vCount) = tTemp(vCount4)
            tDonnee(vCount4, tTemp(4)) = vTemp2
        Next
    Next

' Réaffiche toutes les données triées
    vCount4 = List(0).ListCount - 1
    For vCount = 0 To 3
        List(vCount).Clear
    Next

    For vCount2 = 0 To vCount4
        For vCount = 0 To 3
            List(vCount).AddItem tDonnee(vCount, vCount2)
        Next
    Next
End Sub

Private Sub Form_Activate()
Dim vData As String
Dim vCount As Integer
Dim vCount2 As Integer

    For vCount = 0 To 3
        List(vCount).Clear
    Next
    vCount = 0

    Open "c:\temp\donnees.dat" For Input As #1
    Do
        Line Input #1, vData
        tDonnee(vCount2, vCount) = vData
        List(vCount2).AddItem vData
        vCount2 = vCount2 + 1
        If vCount2 = 4 Then
            tDonnee(4, vCount) = vCount
            vCount2 = 0
            vCount = vCount + 1
        End If
    Loop Until (EOF(1))
    Close #1
End Sub

Private Sub List_Click(Index As Integer)
Dim vCount As Integer

    On Error Resume Next
    For vCount = 0 To 3
        List(vCount).ListIndex = List(Index).ListIndex
    Next
End Sub

Private Sub List_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If List(vCount).ListIndex = -1 Then
            MsgBox "Veuillez sélectionner quelqu'un", vbCritical, "Erreur"
        Else
            If List(3).Text = "" Then
                MenuFichierMail.Enabled = False
            Else
                MenuFichierMail.Enabled = True
            End If
            PopupMenu MenuFichier
        End If
    End If
End Sub

Private Sub MenuFichierDel_Click()
On Error Resume Next
    fSupprimer tDonnee(4, List(0).ListIndex)
End Sub

Private Sub MenuFichierMail_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & List(3).Text & "?subject=Anniversaire", vbMaximizedFocus
End Sub
