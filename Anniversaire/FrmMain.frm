VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu - Nous sommes le"
   ClientHeight    =   2295
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Rechercher"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "&Consulter"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu MenuFichierAdd 
         Caption         =   "Ajouter qqun"
      End
      Begin VB.Menu MenuFichierDel 
         Caption         =   "Supprimer qqun"
      End
      Begin VB.Menu MenuFichierShow 
         Caption         =   "Consulter la liste"
      End
      Begin VB.Menu Tiret1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu MenuAide 
      Caption         =   "Aide"
      Begin VB.Menu MenuAideAbout 
         Caption         =   "A propos"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
    FrmAdd.Show
    FrmMain.Hide
End Sub

Private Sub CmdDel_Click()
    FrmMain.Hide
    FrmDel.Show
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdSearch_Click()
    FrmSearch.Show
    FrmMain.Hide
End Sub

Private Sub CmdShow_Click()
    FrmMain.Hide
    FrmShow.Show
End Sub

Private Sub Form_Load()
Dim vTest As Boolean
Dim vData(4) As String
Dim vYes As Integer
Dim vCount As Integer

    FrmMain.Caption = "Menu - Nous sommes le " & Date
' Lit dans le fichier si il y a un anniversaire aujourd'hui
    Open "c:\temp\donnees.dat" For Input As #1
    Do
        Line Input #1, vData(vCount)
        If vCount = 2 Then
            If Day(vData(vCount)) = Day(Date) And Month(vData(vCount)) = Month(Date) Then
                vNom = vData(vCount - 1)
                vDate = vData(vCount)
                Line Input #1, vMailAnni
                If vMailAnni = "" Then
                    FrmAnni.CmdMail.Enabled = False
                Else
                    FrmAnni.CmdMail.Enabled = True
                End If
                FrmAnni.Show
                FrmMain.Hide
                Close #1
                Exit Sub
            End If
        End If
        vCount = vCount + 1
        If vCount = 4 Then
            vCount = 0
        End If
    Loop Until (EOF(1))
    If (MsgBox("Il n'y a pas d'anniversaire aujoud'hui" & Chr$(13) & "Voulez-vous quitter ?!?", vbYesNo, "Aujourd'hui")) = vbYes Then
        Close #1
        End
    End If
    Close #1
End Sub

Private Sub MenuAideAbout_Click()
    FrmAbout.Show
    FrmMain.Hide
End Sub

Private Sub MenuFichierAdd_Click()
    CmdAdd_Click
End Sub

Private Sub MenuFichierDel_Click()
    CmdDel_Click
End Sub

Private Sub MenuFichierQuitter_Click()
    End
End Sub

Private Sub MenuFichierShow_Click()
    CmdShow_Click
End Sub
