VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculateur de ressources"
   ClientHeight    =   5370
   ClientLeft      =   225
   ClientTop       =   780
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameTemps 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox TxtTemps 
         Height          =   285
         Index           =   0
         Left            =   480
         MaxLength       =   2
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxtTemps 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxtTemps 
         Height          =   285
         Index           =   2
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxtTemps 
         Height          =   285
         Index           =   3
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Minute"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Heure"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Jour"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Seconde"
         Height          =   255
         Left            =   3840
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   14
      Left            =   3720
      TabIndex        =   29
      Text            =   "0"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   13
      Left            =   2520
      TabIndex        =   28
      Text            =   "0"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   12
      Left            =   1320
      TabIndex        =   27
      Text            =   "0"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Text            =   "0"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   10
      Left            =   2520
      TabIndex        =   10
      Text            =   "0"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   11
      Left            =   3720
      TabIndex        =   11
      Text            =   "0"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Text            =   "0"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   8
      Left            =   3720
      TabIndex        =   8
      Text            =   "0"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "0"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TxtRessource 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton CmdCalculer 
      Caption         =   "Calculer"
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   4800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   4800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   4800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   4800
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Trop plein"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "Total"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Tu veux"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "Tu gagnes (par heure)"
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Tu possèdes"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   4800
      X2              =   4800
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Label Label10 
      Caption         =   "Deutérium"
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Cristal"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Métal"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   240
      Width           =   495
   End
   Begin VB.Label LblResultat 
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu FichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu Calculer 
      Caption         =   "Calculer"
      Begin VB.Menu CalculerTempsRes 
         Caption         =   "Temps pour ressources"
         Checked         =   -1  'True
      End
      Begin VB.Menu CalculerResTemps 
         Caption         =   "Ressource en fonction du temps"
      End
      Begin VB.Menu tiret 
         Caption         =   "-"
      End
      Begin VB.Menu CalculeFaire 
         Caption         =   "Faire le calcul"
      End
   End
   Begin VB.Menu About 
      Caption         =   "A Propos"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function SecondeEnHeure(ByVal Seconde As Long) As String
Dim JJ, HH, MM, MMsec, SS As Long
Dim JJstr, HHstr, MMstr, SSstr As String
    
    JJ = Int(Seconde) \ 86400
    HH = (Int(Seconde) - JJ * 86400) \ 3600
    MM = (Int(Seconde) - JJ * 86400 - HH * 3600) \ 60
    SS = Int(Seconde) - JJ * 86400 - HH * 3600 - MM * 60
    
    If SS < 10 Then
        SSstr = "0" + Trim(Str(SS))
    Else
        SSstr = Trim(Str(SS))
    End If

    If MM < 10 Then
        MMstr = "0" + Trim(Str(MM))
    Else
        MMstr = Trim(Str(MM))
    End If

    If HH < 10 Then
        HHstr = "0" + Trim(Str(HH))
    Else
        HHstr = Trim(Str(HH))
    End If

    JJstr = Trim(JJ)
    SecondeEnHeure = JJstr & " jours " & HHstr & "h " & MMstr & "m " & SSstr & "sec "
End Function

Private Sub About_Click()
    FrmAbout.Show
End Sub

Private Sub CalculeFaire_Click()
    CmdCalculer_Click
End Sub

Private Sub CalculerResTemps_Click()
    CalculerTempsRes.Checked = False
    CalculerResTemps.Checked = True
    FrameTemps.Visible = True
End Sub

Private Sub CalculerTempsRes_Click()
    CalculerTempsRes.Checked = True
    CalculerResTemps.Checked = False
    FrameTemps.Visible = False
End Sub

Private Sub CmdCalculer_Click()
Dim vCount As Integer
Dim vManque As Double
Dim vTemps As Variant

    If CalculerTempsRes.Checked = True Then
        For vCount = 0 To 8
            If TxtRessource(vCount).Text = "" Then
                MsgBox "Il manque des paramètres", vbCritical, "Erreur"
                Exit Sub
            End If
        Next
    
        vManque = Int(TxtRessource(6).Text) - Int(TxtRessource(0).Text)
        If Int(TxtRessource(3).Text) <> 0 Then
            vTemps = vManque / Int(TxtRessource(3).Text)
        End If
        
        vManque = Int(TxtRessource(7).Text) - Int(TxtRessource(1).Text)
        If Int(TxtRessource(4).Text) <> 0 Then
            If vManque / Int(TxtRessource(4).Text) > vTemps Then
                vTemps = vManque / Int(TxtRessource(4).Text)
            End If
        End If
    
        vManque = Int(TxtRessource(8).Text) - Int(TxtRessource(2).Text)
        If Int(TxtRessource(5).Text) <> 0 Then
            If vManque / Int(TxtRessource(5).Text) > vTemps Then
                vTemps = vManque / Int(TxtRessource(5).Text)
            End If
        End If

        If (vTemps < 0) Then
            vTemps = 0
        End If
        LblResultat.Caption = "Temps restant : " & SecondeEnHeure(vTemps * 3600)
        TxtRessource(9).Text = Int(Int(TxtRessource(3).Text) * vTemps + Int(TxtRessource(0).Text))
        TxtRessource(10).Text = Int(Int(TxtRessource(4).Text) * vTemps + Int(TxtRessource(1).Text))
        TxtRessource(11).Text = Int(Int(TxtRessource(5).Text) * vTemps + Int(TxtRessource(2).Text))
    
        TxtRessource(12).Text = Int(Int(TxtRessource(9).Text) - Int(TxtRessource(6).Text))
        TxtRessource(13).Text = Int(Int(TxtRessource(10).Text) - Int(TxtRessource(7).Text))
        TxtRessource(14).Text = Int(Int(TxtRessource(11).Text) - Int(TxtRessource(8).Text))
    
    Else    ' En fonction du temps entré
        For vCount = 0 To 5
            If TxtRessource(vCount).Text = "" Then
                MsgBox "Il manque des paramètres", vbCritical, "Erreur"
                Exit Sub
            End If
        Next
        For vCount = 0 To 3
            If TxtTemps(vCount).Text = "" Then
                MsgBox "Il manque des paramètres", vbCritical, "Erreur"
                Exit Sub
            End If
        Next

        TxtRessource(9).Text = Int((Int(TxtTemps(0).Text) * 24 * 3600 + Int(TxtTemps(1).Text) * 3600 + Int(TxtTemps(2).Text) * 60 + Int(TxtTemps(3).Text)) / 3600 * Int(TxtRessource(3).Text) + Int(TxtRessource(0).Text))
        TxtRessource(10).Text = Int((Int(TxtTemps(0).Text) * 24 * 3600 + Int(TxtTemps(1).Text) * 3600 + Int(TxtTemps(2).Text) * 60 + Int(TxtTemps(3).Text)) / 3600 * Int(TxtRessource(4).Text) + Int(TxtRessource(1).Text))
        TxtRessource(11).Text = Int((Int(TxtTemps(0).Text) * 24 * 3600 + Int(TxtTemps(1).Text) * 3600 + Int(TxtTemps(2).Text) * 60 + Int(TxtTemps(3).Text)) / 3600 * Int(TxtRessource(5).Text) + Int(TxtRessource(2).Text))
    End If

    'Sauvegardes des gains
    Open "Gain.dat" For Output As #1
    For vCount = 0 To 8
        vText = TxtRessource(vCount).Text
        Print #1, vText
    Next
    Close #1
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub FichierQuitter_Click()
    End
End Sub

Private Sub Form_Load()
Dim vText As String
Dim vCount As Integer

    On Error Resume Next
    Open "Gain.dat" For Input As #1
    For vCount = 0 To 8
        Line Input #1, vText
        TxtRessource(vCount).Text = vText
    Next
    Close #1
End Sub

Private Sub TxtRessource_GotFocus(Index As Integer)
    TxtRessource(Index).SelStart = 0
    TxtRessource(Index).SelLength = Len(TxtRessource(Index).Text)
End Sub

Private Sub TxtRessource_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCalculer_Click
    ElseIf KeyAscii = 27 Then
        End
    ElseIf Chr(KeyAscii) = "k" Then
        TxtRessource(Index).Text = TxtRessource(Index).Text & "000"
        If Index > 7 Then
            Index = -1
        End If
        TxtRessource(Index + 1).SetFocus
        KeyAscii = 0
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        MsgBox "N'entrez que des nombres", vbCritical, "Erreur"
    End If
End Sub

Private Sub TxtTemps_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCalculer_Click
    ElseIf KeyAscii = 27 Then
        End
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        MsgBox "N'entrez que des nombres", vbCritical, "Erreur"
    End If
End Sub
