VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mélange Lettres"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox TxtOut 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox TxtIn 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vMot As String

Private Sub CmdClear_Click()
    TxtIn.Text = ""
    TxtOut.Text = ""
    vMot = ""
End Sub

Private Sub CmdQuitter_Click()
End
End Sub

Private Sub TxtIn_KeyPress(KeyAscii As Integer)
Dim vMelange As String
Dim tDejaPris(50) As Boolean
Dim vCount As Integer
Dim vChoix As Integer
Dim vPlus As String

    If Chr(KeyAscii) = " " Then
        If Len(vMot) > 3 Then
            vMelange = Left(vMot, 1)
            For vCount = 0 To Len(vMot) - 3
                Do
                    vChoix = Int(Rnd * (Len(vMot) - 2) + 1)
                    If tDejaPris(vChoix) = 0 Then
                        tDejaPris(vChoix) = True
                        vMelange = vMelange & Mid(vMot, vChoix + 1, 1)
                        vChoix = 0
                    End If
                Loop While (vChoix <> 0)
            Next
            vMelange = vMelange & Right(vMot, 1)
            vMot = vMelange
        End If
        If TxtOut.Text = "" Then
            TxtOut.Text = vMot
        Else
            TxtOut.Text = TxtOut.Text & Chr(KeyAscii) & vMot
        End If
        vMot = ""
    ElseIf Chr(KeyAscii) = "!" Or Chr(KeyAscii) = ":" Or Chr(KeyAscii) = "," Or Chr(KeyAscii) = "?" Or Chr(KeyAscii) = "'" Or IsNumeric(Chr(KeyAscii)) Then
        If Right(TxtOut.Text, 1) <> " " Then
            vPlus = " "
        End If
        TxtOut.Text = TxtOut.Text & vPlus & vMot & Chr(KeyAscii)
        vMot = ""
    ElseIf KeyAscii = 8 Then
        If vMot <> "" Then
            vMot = Left(vMot, Len(vMot) - 1)
        End If
    Else
        vMot = vMot + Chr(KeyAscii)
    End If
End Sub
