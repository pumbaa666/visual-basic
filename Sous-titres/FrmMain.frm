VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Modifier le temps pour les sous-titres"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   3360
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox TxtSec 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Text            =   "Sec"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox TxtMin 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "Min"
      Top             =   3960
      Width           =   375
   End
   Begin VB.OptionButton OptMoins 
      Caption         =   "En moins"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.OptionButton OptPlus 
      Caption         =   "En plus"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ComboBox CmbExt 
      Height          =   315
      ItemData        =   "FrmMain.frx":0000
      Left            =   3360
      List            =   "FrmMain.frx":000A
      TabIndex        =   4
      Text            =   "Type"
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Combien de temps"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Choisissez le fichier à modifier"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbExt_Click()
    File1.Pattern = CmbExt.Text
End Sub

Private Sub CmdOk_Click()
Dim vSec As Integer
Dim vMin As Integer
Dim vEnCours As String
Dim vFichier As String
Dim vNewSec As String
Dim vNewMin As String
Dim vNewHeure As String
Dim vNewTime As String

    If TxtMin.Text = "" Or TxtMin.Text = "Min" Then
        vMin = 0
    Else
        vMin = TxtMin.Text
    End If
    
    If TxtSec.Text = "" Or TxtSec.Text = "Sec" Then
        vSec = 0
    Else
        vSec = TxtSec.Text
    End If

    If vMin > 59 Or vSec > 59 Or (vMin = 0 And vSec = 0) Then
        MsgBox "Valeurs incorrectes", vbCritical, "Erreur"
    ElseIf File1.FileName = "" Then
        MsgBox "Choisissez un fichier", vbCritical, "Erreur"
    Else
        If Len(File1.Path) = 3 Then
            vFichier = File1.Path & File1.FileName
        Else
            vFichier = File1.Path & "\" & File1.FileName
        End If
        Open vFichier For Input As #1
        Open "c:\New.tmp" For Output As #2
        Do
            Line Input #1, vEnCours
            If Len(vEnCours) = 29 Then
                If IsDate(Left(vEnCours, 8)) Then
                    '********** 1ère heure **********'
                    vNewSec = Mid(vEnCours, 7, 2) + vSec
                    If vNewSec < 10 Then
                        vNewSec = "0" & vSec
                    End If

                    vNewMin = Mid(vEnCours, 4, 2) + vMin
                    If vNewMin < 10 Then vNewMin = "0" & vNewMin

                    vNewHeure = Left(vEnCours, 2)
                    If vNewSec > 59 Then
                        vNewSec = vNewSec - 60
                        vNewMin = vNewMin + 1
                    End If
                    vNewTime1 = vNewHeure & ":" & vNewMin & ":" & vNewSec & Mid(vEnCours, 9, 4)
                    '**********************************'
                    
                    '********** 2ème heure **********'
                    vNewSec = Mid(vEnCours, 24, 2) + vSec
                    If vNewSec < 10 Then vNewSec = "0" & vSec

                    vNewMin = Mid(vEnCours, 21, 2) + vMin
                    If vNewMin < 10 Then vNewMin = "0" & vNewMin

                    vNewHeure = Mid(vEnCours, 18, 2)
                    If vNewSec > 59 Then
                        vNewSec = vNewSec - 60
                        If vNewSec < 10 Then vNewSec = "0" & vNewSec

                        vNewMin = vNewMin + 1
                        If vNewMin < 10 Then vNewMin = "0" & vNewMin
                    End If
                    vNewTime2 = vNewHeure & ":" & vNewMin & ":" & vNewSec & Right(vEnCours, 4)
                    '**********************************'
                    vEnCours = vNewTime1 & " --> " & vNewTime2
                End If
            End If
            Print #2, vEnCours
        Loop While (EOF(1) = False)
        Close #2
        Close #1
        Kill vFichier
        Name "c:\New.tmp" As vFichier
        MsgBox "Remplacement terminé", vbInformation, "Fini"
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    CmdOk_Click
End Sub

Private Sub TxtMin_GotFocus()
    If TxtMin.Text = "Min" Then TxtMin.Text = ""
End Sub

Private Sub TxtSec_GotFocus()
    If TxtSec.Text = "Sec" Then TxtSec.Text = ""
End Sub
