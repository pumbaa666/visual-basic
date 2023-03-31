VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion d'équipement"
   ClientHeight    =   9375
   ClientLeft      =   225
   ClientTop       =   780
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCouleur 
      Height          =   1095
      Index           =   10
      Left            =   360
      Picture         =   "FrmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "301"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton CmdCouleur 
      Height          =   1095
      Index           =   9
      Left            =   2760
      Picture         =   "FrmMain.frx":10EF
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "300"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Frame FrameBoutton 
      Caption         =   "Afficher les outils"
      Height          =   9015
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   17
         Left            =   2640
         Picture         =   "FrmMain.frx":2264
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "360"
         Top             =   7680
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   16
         Left            =   240
         Picture         =   "FrmMain.frx":3496
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "331"
         Top             =   7680
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   15
         Left            =   2640
         Picture         =   "FrmMain.frx":439A
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "320"
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   14
         Left            =   240
         Picture         =   "FrmMain.frx":5576
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "315"
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   13
         Left            =   2640
         Picture         =   "FrmMain.frx":691B
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "312"
         Top             =   5280
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   12
         Left            =   240
         Picture         =   "FrmMain.frx":7CDF
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "306"
         Top             =   5280
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   11
         Left            =   2640
         Picture         =   "FrmMain.frx":8E2B
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "305"
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Height          =   1095
         Index           =   8
         Left            =   240
         Picture         =   "FrmMain.frx":A220
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "005"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Rouge"
         Height          =   495
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         Tag             =   "Rouge"
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Vert"
         Height          =   495
         Index           =   1
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "Vert"
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Bleu"
         Height          =   495
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Bleu"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Local/Réserve"
         Height          =   495
         Index           =   3
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Local/Réserve"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Révision"
         Height          =   495
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Révision"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Défectueux"
         Height          =   495
         Index           =   5
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Défectueux"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "En prêt"
         Height          =   495
         Index           =   6
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "En prêt"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton CmdCouleur 
         Caption         =   "Cours/Exercice"
         Height          =   495
         Index           =   7
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Cours/Exercice"
         Top             =   2280
         Width           =   2175
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Outils"
      Height          =   1335
      Left            =   5400
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton Cmd 
         Caption         =   "1"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.FileListBox File 
      Height          =   2235
      Left            =   720
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   615
      Left            =   5400
      TabIndex        =   18
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu FichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu Aide 
      Caption         =   "Aide"
      Begin VB.Menu AideUtilisation 
         Caption         =   "Utilisation"
      End
      Begin VB.Menu AideApropos 
         Caption         =   "A propos"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vFrameEnCours As Integer
Dim vInit As Boolean
Dim tTab(1000, 3) As String
Dim vNbElem As Integer

Private Sub AideApropos_Click()
    FrmAPropos.Show
End Sub

Private Sub AideUtilisation_Click()
    FrmAide.Show
End Sub

Private Sub Cmd_Click(Index As Integer)
    tEnCours(0) = tTab(Cmd(Index).Tag, 0)
    tEnCours(1) = tTab(Cmd(Index).Tag, 1)
    tEnCours(2) = tTab(Cmd(Index).Tag, 2)
    tEnCours(3) = Cmd(Index).Tag
    FrmChanger.Show
End Sub

Private Sub CmdCouleur_Click(Index As Integer)
Dim iRouge As Integer
Dim i As Integer
Dim j As Integer
Dim vNbY As Integer
Dim vMaxYRouge As Integer
Dim vParametre As Integer
Static vNbCmd As Integer

    LectureFichier
    If Index <= 2 Then
        vParametre = 0
    Else
        vParametre = 1
    End If

    vNbY = 15
    vMaxYRouge = Cmd(0).Top
    Do
        If LCase(tTab(i, 1 + vParametre)) = LCase(CmdCouleur(Index).Tag) Then
            If iRouge = 0 Then
                If LCase(tTab(i, 1)) = "rouge" Then
                    Cmd(iRouge).BackColor = &HFF&
                ElseIf LCase(tTab(i, 1)) = "bleu" Then
                    Cmd(iRouge).BackColor = &HFFFF00
                ElseIf LCase(tTab(i, 1)) = "vert" Then
                    Cmd(iRouge).BackColor = &HFF00&
                End If
                Cmd(iRouge).Caption = tTab(i, 0)
            Else
                If iRouge >= vNbCmd Then
                    Load Cmd(iRouge)
                End If
                Cmd(iRouge).Caption = tTab(i, 0)
                If LCase(tTab(i, 1)) = "rouge" Then
                    Cmd(iRouge).BackColor = &HFF&
                ElseIf LCase(tTab(i, 1)) = "bleu" Then
                    Cmd(iRouge).BackColor = &HFFFF00
                ElseIf LCase(tTab(i, 1)) = "vert" Then
                    Cmd(iRouge).BackColor = &HFF00&
                End If
                Cmd(iRouge).Top = Cmd(iRouge - 1).Top + Cmd(iRouge - 1).Height
                Cmd(iRouge).Left = Cmd(iRouge - 1).Left
                If iRouge Mod vNbY = 0 Then
                    Cmd(iRouge).Top = Cmd(0).Top
                    Cmd(iRouge).Left = Cmd(iRouge - 1).Left + Cmd(0).Width + 100
                ElseIf Cmd(iRouge).Top > vMaxYRouge Then
                    vMaxYRouge = Cmd(iRouge).Top
                End If
            End If
            Cmd(iRouge).Tag = i
            Cmd(iRouge).ToolTipText = tTab(i, 2)
            Cmd(iRouge).Visible = True
            iRouge = iRouge + 1
        End If
        i = i + 1
    Loop While (tTab(i, 0) <> "")

    If iRouge = 0 Then
        iRouge = 1
        Frame.Visible = False
'        FrmMain.Width = CmdQuitter.Left + CmdQuitter.Width + 500
    Else
        Frame.Visible = True
        Frame.Height = vMaxYRouge + 500
        Frame.Width = Cmd(iRouge - 1).Left + Cmd(iRouge - 1).Width + 150
        FrmMain.Width = Frame.Left + Frame.Width + 500
    End If
    
    For j = iRouge To vNbCmd - 1
        Unload Cmd(j)
    Next
    vNbCmd = iRouge
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Function LectureFichier()
Dim vString As String
Dim i As Integer

    ViderTableau
    Open CheminRelatif & "\pss 90.txt" For Input As #1
    i = 0
    Do
        Line Input #1, vString
        tTab(i, 0) = Left(vString, 9)
        tTab(i, 1) = Trim(Mid(vString, 11, 5))
        tTab(i, 2) = Right(vString, Len(vString) - 16)
        i = i + 1
    Loop Until (EOF(1))
    Close 1
    vNbElem = i - 1
End Function

Private Function ViderTableau()
Dim i As Integer

    i = 0
    Do Until (tTab(i, 0) = "")
        tTab(i, 0) = ""
        tTab(i, 1) = ""
        tTab(i, 2) = ""
        i = i + 1
    Loop
End Function

Private Sub FichierQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    CheminRelatif = App.Path
    If Right(CheminRelatif, 1) <> "\" Then
        CheminRelatif = CheminRelatif & "\"
    End If
End Sub

