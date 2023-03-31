VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Séries"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdHide 
      Caption         =   "&Masquer"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   480
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

Private Sub CmdHide_Click()
    FrmMain.Hide
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    On Error GoTo CreatFile
    Open "c:\temp\nbseries.dat" For Input As #1
    Line Input #1, vNbEnreg
    Close #1
    Exit Sub
      
CreatFile:
    Close #1
    Open "c:\temp\nbseries.dat" For Output As #1
    Print #1, "1"
    Close #1
    vNbEnreg = "1"
End Sub

Private Sub Timer1_Timer()
Static vCount As Integer
Static vCount2 As Integer
Static sMain As Structure
Dim vJour As Integer

    For vCount = 1 To vNbEnreg - 1
        Open "c:\temp\series.dat" For Random As #1 Len = Len(sMain)
        Get #1, vCount, sMain
        Close #1
        
        If Format(Date, "dddd") = "lundi" Then
            vJour = 0
        ElseIf Format(Date, "dddd") = "mardi" Then
            vJour = 1
        ElseIf Format(Date, "dddd") = "mercredi" Then
            vJour = 2
        ElseIf Format(Date, "dddd") = "jeudi" Then
            vJour = 3
        ElseIf Format(Date, "dddd") = "vendredi" Then
            vJour = 4
        ElseIf Format(Date, "dddd") = "samedi" Then
            vJour = 5
        ElseIf Format(Date, "dddd") = "dimanche" Then
            vJour = 6
        End If
        
        If sMain.vJours(vJour) = 1 And Time = sMain.vHeure Then
            MsgBox sMain.vTitre, vbCritical, "Vite"
        End If
    Next
    
'    If Format(Date, "dddd") <> "samedi" And Format(Date, "dddd") <> "dimanche" Then
'        If Time = "17:55:00" Then
'            MsgBox "StarGate", vbCritical, "Vite!!!"
'        ElseIf Time = "20:00:00" Then
'            MsgBox "Une Nounou d'enfer", vbCritical, "Vite!!!"
'            End
'        End If
'    ElseIf Format(Date, "dddd") = "dimanche" Then
'        If Time = "10:55:00" Then
'            MsgBox "Grand écran", vbCritical, "Vite!!!"
'        End If
'    End If
End Sub
