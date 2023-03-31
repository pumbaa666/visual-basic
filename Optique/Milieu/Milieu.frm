VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   " Milieu optique"
   ClientHeight    =   6705
   ClientLeft      =   3900
   ClientTop       =   390
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkCreate 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   4080
   End
   Begin VB.Frame FrameInterface 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   6240
      TabIndex        =   4
      Top             =   360
      Width           =   3255
      Begin VB.TextBox TxtX 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtN 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdQuitter 
         Caption         =   "&Quitter"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox TxtAngle 
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Dessiner"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox TxtNbN 
         Height          =   285
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.Label LblA 
         Caption         =   "A1 :"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LblX 
         Caption         =   "X1 :"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblN 
         Caption         =   "N1 :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Angle d'incidence : "
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de milieu : "
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Line LigneRayon 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LigneNormale 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   1320
      X2              =   1320
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Line LigneInterface 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   960
      X2              =   2880
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tAngle(1 To 10) As Double
Dim vNbOk As Integer

Private Sub ClkCreate_Timer()
Static vCount As Boolean
Static vLastNb As Integer

    If vCount = False Then
        vCount = True
    Else
        On Error Resume Next
        If IsNumeric(TxtNbN.Text) Then
            For i = 1 To Int(TxtNbN.Text)
                If Not i < vLastNb + 1 Then
                    Load LblN(i)
                    Load TxtN(i)
                    Load LblX(i)
                    Load TxtX(i)
                    Load LblA(i)
                    Load LigneInterface(i)
                    Load LigneNormale(i)
                    Load LigneRayon(i)
                End If
                LblN(i).Caption = "N" & i & " :"
                LblN(i).Top = LblN(i - 1).Top + 360
                LblN(i).Visible = True

                TxtN(i).Top = LblN(i - 1).Top + 360
                TxtN(i).TabIndex = 2 * i + 2
                TxtN(i).Visible = True
                
                LblX(i).Caption = "X" & i & " :"
                LblX(i).Top = LblN(i - 1).Top + 360
                LblX(i).Visible = True

                TxtX(i).Top = LblN(i - 1).Top + 360
                TxtX(i).TabIndex = 2 * i + 3
                TxtX(i).Visible = True
            
                LblA(i).Caption = "A" & i & " :"
                LblA(i).Top = LblN(i - 1).Top + 360
                LblA(i).Visible = True

                LigneRayon(i).Visible = True
            Next

            If vLastNb >= i Then
                For j = i To vLastNb
                    Unload LblN(j)
                    Unload TxtN(j)
                    Unload LblX(j)
                    Unload TxtX(j)
                    Unload LblA(j)
                    Unload LigneInterface(j)
                    Unload LigneNormale(j)
                    Unload LigneRayon(j)
                Next
            End If

            vLastNb = i - 1
            CmdOk.Top = LblN(i - 1).Top + 500
            CmdOk.TabIndex = 2 * i + 3
            CmdQuitter.Top = CmdOk.Top + 700
            CmdQuitter.TabIndex = 2 * i + 4
            FrameInterface.Height = CmdQuitter.Top + CmdQuitter.Height + 360
        End If
        ClkCreate.Enabled = False
    End If
End Sub

Private Sub CmdOk_Click()
    If Valeur = True Then
        vNbOk = CalculAngles
        DessinerInterface
    End If
End Sub

Private Function CalculAngles() As Integer
    tAngle(1) = TxtAngle.Text
    LblA(1).Caption = "A1: " & Left(tAngle(1), 4)
    For i = 2 To Int(TxtNbN.Text)
        tAngle(i) = TxtN(i - 1).Text / TxtN(i).Text * Sin(tAngle(i - 1) * 3.14159 / 180)
        If (tAngle(i) > 1) Then
            LblA(i).Caption = "R T"
            Exit For
        Else
            tAngle(i) = Atn(tAngle(i) / Sqr(-tAngle(i) * tAngle(i) + 1))
            tAngle(i) = tAngle(i) * 180 / 3.14159
            LblA(i).Caption = "A" & i & ": " & Left(tAngle(i), 4)
        End If
    Next
    CalculAngles = i - 1
End Function

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    LigneInterface(0).X1 = 0
    LigneInterface(0).Y1 = 0
    LigneInterface(0).X2 = 0
    LigneInterface(0).Y2 = 0
    LigneRayon(0).X1 = 0
    LigneRayon(0).Y1 = 0
    LigneRayon(0).X2 = 0
    LigneRayon(0).Y2 = 0
End Sub

Private Sub Form_Resize()
    FrameInterface.Left = FrmMain.Width - FrameInterface.Width - 100
    DessinerInterface
End Sub

Private Function DessinerInterface()
Dim vTaille As Double

    If IsNumeric(TxtNbN.Text) Then
        For i = 1 To Int(TxtNbN.Text)
            vTaille = vTaille + TxtX(i).Text
        Next
        LigneRayon(0).X1 = 500
        LigneRayon(0).Y1 = 0
        For i = 1 To Int(TxtNbN.Text)
            LigneInterface(i).X1 = 0
            LigneInterface(i).Y1 = TxtX(i).Text / vTaille * FrmMain.Height + LigneInterface(i - 1).Y1
            LigneInterface(i).X2 = FrameInterface.Left - 200
            LigneInterface(i).Y2 = LigneInterface(i).Y1
            LigneInterface(i).Visible = True

            LigneRayon(i - 1).X2 = (LigneInterface(i).Y1 - LigneInterface(i - 1).Y1) / Abs(Tan(tAngle(i) * 3.14159 / 180)) + LigneRayon(i - 1).X1
            LigneRayon(i - 1).Y2 = LigneInterface(i).Y1
            LigneRayon(i).X1 = LigneRayon(i - 1).X2
            LigneRayon(i).Y1 = LigneRayon(i - 1).Y2

            LigneNormale(i).X1 = LigneRayon(i).X1
            LigneNormale(i).Y1 = LigneInterface(i).Y1 - 500
            LigneNormale(i).X2 = LigneNormale(i).X1
            LigneNormale(i).Y2 = LigneInterface(i).Y1 + 500
            LigneNormale(i).Visible = True
        Next
        i = i - 1
        LigneRayon(i).X2 = (LigneInterface(i).Y1 - LigneInterface(i - 1).Y1) / -Tan(tAngle(i))
        LigneRayon(i).Y2 = LigneInterface(i).Y1

        FrmMain.Width = LigneRayon(i).X1 + FrameInterface.Width + 500
    End If
End Function

Private Sub TxtNbN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    Else
        ClkCreate.Enabled = True
    End If
End Sub

Private Sub TxtAngle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Function Valeur() As Boolean
    Valeur = True
    If IsNumeric(TxtNbN.Text) = False Or IsNumeric(TxtAngle.Text) = False Then
        MsgBox "Il faut entrer des valeurs numériques", vbCritical, "Erreur"
        Valeur = False
    ElseIf Int(TxtNbN.Text) < 2 Then
        MsgBox "Le nombre de milieu doit être compris entre 2 et 9", vbCritical, "Erreur"
        Valeur = False
    ElseIf Int(TxtAngle.Text) < 1 Or Int(TxtAngle.Text) > 90 Then
        MsgBox "L'angle doit être compris entre 1 et 90°", vbCritical, "Erreur"
        Valeur = False
    Else
        For i = 1 To Int(TxtNbN.Text)
            If IsNumeric(TxtN(i).Text) = False Or IsNumeric(TxtX(i).Text) = False Then
                MsgBox "Valeur de n/x " & i & " incorrect", vbCritical, "Erreur"
                Valeur = False
                Exit For
            ElseIf Int(TxtN(i).Text) < 1 Or Int(TxtN(i).Text) <= 0 Then
                MsgBox "Valeur de n/x " & i & " incorrect", vbCritical, "Erreur"
                Valeur = False
                Exit For
            End If
        Next
    End If
End Function

Private Sub TxtX_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub

Private Sub TxtN_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOk_Click
    End If
End Sub
