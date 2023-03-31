VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mes DVD's - Menu"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Fin"
      Height          =   1575
      Left            =   2280
      TabIndex        =   13
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton CmdQuitter 
         Caption         =   "&Quitter"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Sauver"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Timer ClkLoad 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Localisation"
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
      Begin VB.CommandButton CmdAtteindre 
         Caption         =   "A&tteindre"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton CmdChercher 
         Caption         =   "&Chercher"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sorties"
      Height          =   2175
      Left            =   2280
      TabIndex        =   11
      Top             =   240
      Width           =   1695
      Begin VB.CommandButton CmdNbPreter 
         Caption         =   "&Nb Prêtés"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdPreter 
         Caption         =   "&Prêter"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton CmdReprendre 
         Caption         =   "&Reprendre"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modifications"
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1815
      Begin VB.CommandButton CmdModifier 
         Caption         =   "&Modifier"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Ajouter"
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "&Supprimer"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClkLoad_Timer()
Static vCntLoad As Boolean
    If vCntLoad = False Then
        FrmWait.Show
        vCntLoad = True
    Else
        Setup
        LectureFichier

        FrmMain.Top = 0
        FrmMain.Left = 500
        FrmMain.Show

        FrmListe.Top = 4900
        FrmListe.Left = 500
        FrmListe.Show

        FrmWait.Hide
        ClkLoad.Enabled = False

        FrmMain.SetFocus

        FrmWait.Label1.Caption = "Suppression en cours, veuillez patienter"
    End If
End Sub

Private Sub CmdAdd_Click()
    ToutCacher
    FrmAdd.Top = 100
    FrmAdd.Left = 7000
    FrmAdd.Caption = "Mes DVD - Ajouter un DVD"
    FrmAdd.Height = 3105
    FrmAdd.Frame1.Top = 240
    FrmAdd.TxtNum.Visible = False
    FrmAdd.Show
End Sub

Private Sub CmdAtteindre_Click()
    ToutCacher
    CmdReprendre_Click
    FrmPreter.Caption = "Mes DVD - Atteindre un DVD"
End Sub

Private Sub CmdChercher_Click()
    ToutCacher
    FrmChercher.Top = 100
    FrmChercher.Left = 7000
    FrmChercher.Show
End Sub

Private Sub CmdDel_Click()
    CmdReprendre_Click
    FrmPreter.Caption = "Mes DVD - Supprimer un DVD"
End Sub

Private Sub CmdModifier_Click()
    BouttonModifier
End Sub

Private Sub CmdNbPreter_Click()
Dim vCount As Integer
Dim vNbPrete As Integer
    RefreshListe
    For vCount = 0 To FrmListe.Liste(5).ListCount - 1
        FrmListe.Liste(5).ListIndex = vCount
        If FrmListe.Liste(5).Text <> "" Then
            vNbPrete = vNbPrete + 1
        End If
    Next
    CmdNbPreter.Caption = "Nb prêtés : " & vNbPrete
    FrmListe.Liste(0).ListIndex = 0
End Sub

Private Sub CmdPreter_Click()
    BouttonPreter
End Sub

Private Sub CmdQuitter_Click()
    CleanUp
    End
End Sub

Private Sub CmdReprendre_Click()
    BouttonReprendre
End Sub

Private Sub CmdSave_Click()
    appDVD.SaveWorkspace
End Sub

Private Sub Form_Activate()
RefreshListe
End Sub

Private Sub Form_Load()
    vTestListe5 = True
End Sub
