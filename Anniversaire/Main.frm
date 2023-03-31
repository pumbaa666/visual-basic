VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anniversaire"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdShow 
      Caption         =   "&Consulter"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Anniversaire"
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
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1695
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

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
' Crée le fichier si il n'existe pas
Dim vTest As Boolean
Dim vData(4) As String
Dim vYes As Integer
Dim vAnni As Boolean
Dim vCount As Integer
    On Error GoTo CreatFile
    Open "c:\temp\donnees.dat" For Input As #1
    Do
        Line Input #1, vData(vCount)
        vTest = 1
        If vData(vCount) = Date Then
            MsgBox "Aie!!! C'est l'anniversaire de " & vData(0) & " aujourd'hui!", vbCritical
            vAnni = 1
        End If
        vCount = vCount + 1
        If vCount = 4 Then
            vCount = 0
        End If
    Loop Until (EOF(1))
    If vAnni = 0 Then
        MsgBox "Il n'y a pas d'anniversaire aujoud'hui", vbOKOnly
        vYes = MsgBox("Voulez-vous quitter ?!?", vbYesNo)
        If vYes = vbYes Then
            End
        End If
    End If
    Close #1
CreatFile:
    If vTest = 0 Then
        Close #1
        Open "c:\temp\donnees.dat" For Append As #1
        Close #1
    End If
End Sub
