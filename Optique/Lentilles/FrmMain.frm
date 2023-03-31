VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Optique"
   ClientHeight    =   7710
   ClientLeft      =   6285
   ClientTop       =   615
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameInfo 
      Caption         =   "Informations"
      Height          =   2415
      Left            =   9840
      TabIndex        =   15
      Top             =   3720
      Width           =   1575
      Begin VB.Label LblI 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label LblGamma 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label LblQ 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LblImage 
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblObjet 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdDessiner 
      Caption         =   "&Dessiner"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CheckBox ChkInterface 
      Caption         =   "Afficher Interface"
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Frame FrameValeurs 
      Caption         =   "Valeurs"
      Height          =   1575
      Left            =   9840
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
      Begin VB.TextBox TxtValO 
         Height          =   285
         Left            =   480
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox TxtValF 
         Height          =   285
         Left            =   480
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtValP 
         Height          =   285
         Left            =   480
         MaxLength       =   6
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "O"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "f"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "p"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   9840
      TabIndex        =   8
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Frame FrameType 
      Caption         =   "Type de lentille"
      Height          =   1095
      Left            =   9840
      TabIndex        =   7
      Top             =   480
      Width           =   1575
      Begin VB.OptionButton OptType 
         Caption         =   "Convergente"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptType 
         Caption         =   "Divergente"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label LblNomObjet 
      Caption         =   "Objet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblNomImage 
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line LigneRayon 
      BorderColor     =   &H00C00000&
      Index           =   4
      X1              =   -2160
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LigneRayon 
      BorderColor     =   &H00C00000&
      Index           =   3
      X1              =   -2160
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LigneRayon 
      BorderColor     =   &H00C00000&
      Index           =   2
      X1              =   -2160
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LigneRayon 
      BorderColor     =   &H00C00000&
      Index           =   1
      X1              =   -2160
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LigneRayon 
      BorderColor     =   &H00C00000&
      Index           =   0
      X1              =   -2160
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LigneImage 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -1680
      X2              =   -1680
      Y1              =   5280
      Y2              =   6000
   End
   Begin VB.Line LigneImageFleche1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -1800
      X2              =   -1800
      Y1              =   5280
      Y2              =   5520
   End
   Begin VB.Line LigneImageFleche2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -1560
      X2              =   -1560
      Y1              =   5280
      Y2              =   5520
   End
   Begin VB.Line LigneObjetFleche2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   -1560
      X2              =   -1560
      Y1              =   2760
      Y2              =   3000
   End
   Begin VB.Line LigneObjetFleche1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   -1800
      X2              =   -1800
      Y1              =   2760
      Y2              =   3000
   End
   Begin VB.Line LigneObjet 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   -1680
      X2              =   -1680
      Y1              =   2760
      Y2              =   3480
   End
   Begin VB.Label LblF2 
      Caption         =   "F '"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblF 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line LigneF2 
      Visible         =   0   'False
      X1              =   7560
      X2              =   7560
      Y1              =   3840
      Y2              =   4200
   End
   Begin VB.Line LigneF 
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   3840
      Y2              =   4200
   End
   Begin VB.Line LigneType 
      BorderWidth     =   2
      Index           =   3
      X1              =   5760
      X2              =   6120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line LigneType 
      BorderWidth     =   2
      Index           =   2
      X1              =   5760
      X2              =   6120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line LigneType 
      BorderWidth     =   2
      Index           =   1
      X1              =   5760
      X2              =   6120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line LigneType 
      BorderWidth     =   2
      Index           =   0
      X1              =   5760
      X2              =   6120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line LigneLentille 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   960
      Y2              =   6960
   End
   Begin VB.Line LigneFoyer 
      BorderWidth     =   2
      X1              =   0
      X2              =   11280
      Y1              =   4005
      Y2              =   4005
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vP As Double
Dim vQ As Double
Dim vO As Double
Dim vI As Double
Dim vF As Double
Dim vGamma As Double

Private Function Afficher(vAfficher As Boolean)
Dim vCount As Integer
    For vCount = 0 To 4
        LigneRayon(vCount).Visible = vAfficher
    Next
    LigneObjet.Visible = vAfficher
    LigneObjetFleche1.Visible = vAfficher
    LigneObjetFleche2.Visible = vAfficher
    LigneImage.Visible = vAfficher
    LigneImageFleche1.Visible = vAfficher
    LigneImageFleche2.Visible = vAfficher
    
End Function

Private Function Calcul() As Boolean
    Calcul = False
    If Valeur = False Then
        MsgBox "Valeurs erronées", vbCritical, "Erreur"
    ElseIf Int(TxtValF.Text) < 0 Or Int(TxtValP.Text) < 0 Then
        MsgBox "p et f ne peuvent pas être négatifs", vbCritical, "Erreur"
    ElseIf Int(TxtValF.Text) = 0 Or Int(TxtValP.Text) = 0 Then
        MsgBox "p et f ne peuvent pas être nuls", vbCritical, "Erreur"
    Else
        vP = TxtValP.Text
        vO = TxtValO.Text
        vF = TxtValF.Text
        If OptType(1).Value = True Then
            vF = -1 * vF
        End If
        If vF <> vP Then
            vQ = 1 / (1 / vF - 1 / vP)
            vGamma = -1 * vQ / vP
            vI = vGamma * vO
            LblQ.Caption = "q = " + Left(Str(vQ), 5)
            LblGamma.Caption = "y = " + Left(Str(vGamma), 5)
            LblI.Caption = "I = " + Left(Str(vI), 5)
            Afficher (True)
            Calcul = True
        Else
            LblQ.Caption = "q est à l'infini"
            LblGamma.Caption = ""
            LblI.Caption = ""
            Afficher (False)
        End If
        If vP > 0 Then
            LblObjet.Caption = "L'objet est réel"
        Else
            LblObjet.Caption = "L'objet est virtuel"
        End If
    
        If vQ > 0 Then
            LblImage.Caption = "L'image est réelle"
        Else
            LblImage.Caption = "L'image est virtuelle"
        End If
    
        If vGamma > 0 Then
            LblImage.Caption = LblImage.Caption + " et droite"
        Else
            LblImage.Caption = LblImage.Caption + " et inversée"
        End If
    End If
End Function


Private Sub ChkInterface_Click()
    DessinFond
End Sub

Private Sub CmdDessiner_Click()
    DessinFond
    If Calcul <> False Then
        DessinObjet
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub Form_Load()
    DessinFond
End Sub

Private Function DessinObjet()
    '**************** calcul du ratio largeur ****************'
    Dim vXFinI As Double
    Dim vRatio As Double
    Dim vXF As Integer
    Dim vCount As Integer

    Dim vGrand As Integer
    vXF = Int(TxtValF.Text)
    If Abs(vQ) > Abs(vP) Then
        vGrand = vQ
    Else
        vGrand = vP
    End If
    If Abs(vXF) > Abs(vGrand) Then
        vGrand = vXF
    End If
    vXFinI = vGrand
    Do While Abs(vXFinI) >= 1
        vXFinI = vXFinI / 10
    Loop
    vXFinI = vXFinI * FrmMain.Width / 2
    vRatio = vXFinI / vGrand
    vQ = vQ * vRatio
    vP = vP * vRatio
    vXF = vXF * vRatio

    '**************** Foyer ****************'
    LigneF.X1 = LigneLentille.X1 - vXF
    LigneF.X2 = LigneLentille.X1 - vXF
    LigneF.Y1 = LigneFoyer.Y1 - 150
    LigneF.Y2 = LigneFoyer.Y1 + 150
    LigneF.Visible = True

    LigneF2.X1 = LigneLentille.X1 + vXF
    LigneF2.X2 = LigneLentille.X1 + vXF
    LigneF2.Y1 = LigneFoyer.Y1 - 150
    LigneF2.Y2 = LigneFoyer.Y1 + 150
    LigneF2.Visible = True

    LblF.Top = LigneF.Y2 + 100
    LblF.Left = LigneF.X1 - 50
    LblF.Visible = True

    LblF2.Top = LigneF2.Y2 + 100
    LblF2.Left = LigneF2.X1 - 50
    LblF2.Visible = True

    '**************** calcul du ratio hauteur ****************'
    If Abs(vO) > Abs(vI) Then
        vGrand = vO
    Else
        vGrand = vI
    End If
    vXFinI = vGrand
    Do While Abs(vXFinI) >= 1
        vXFinI = vXFinI / 10
    Loop
    vXFinI = vXFinI * FrmMain.Height / 2
    vRatio = vXFinI / vGrand
    vI = vI * vRatio
    vO = vO * vRatio

    '**************** Objet ****************'
    LigneObjet.X1 = LigneLentille.X1 - vP
    LigneObjet.X2 = LigneLentille.X1 - vP
    LigneObjet.Y1 = LigneFoyer.Y1
    LigneObjet.Y2 = LigneFoyer.Y1 - vO

    LigneObjetFleche1.X1 = LigneObjet.X1
    LigneObjetFleche1.X2 = LigneObjet.X1 - 150
    LigneObjetFleche1.Y1 = LigneObjet.Y2
    LigneObjetFleche2.X1 = LigneObjet.X1
    LigneObjetFleche2.X2 = LigneObjet.X1 + 150
    LigneObjetFleche2.Y1 = LigneObjet.Y2

    If vO > 0 Then
        LigneObjetFleche1.Y2 = LigneObjet.Y2 + 150
        LigneObjetFleche2.Y2 = LigneObjet.Y2 + 150
    Else
        LigneObjetFleche1.Y2 = LigneObjet.Y2 - 150
        LigneObjetFleche2.Y2 = LigneObjet.Y2 - 150
    End If
    Dim vPlus As Integer
    If LigneObjet.Y2 > LigneObjet.Y1 Then
        vPlus = 500
    Else
        vPlus = -500
    End If
    LblNomObjet.Left = LigneObjet.X1 - LblNomObjet.Width
    LblNomObjet.Top = LigneObjet.Y2 + vPlus
   ' LblNomObjet.Visible = True

    '**************** Image ****************'
    LigneImage.X1 = LigneLentille.X1 + vQ
    LigneImage.X2 = LigneLentille.X1 + vQ
    LigneImage.Y1 = LigneFoyer.Y1
    LigneImage.Y2 = LigneFoyer.Y1 - vI

    LigneImageFleche1.X1 = LigneImage.X1
    LigneImageFleche1.X2 = LigneImage.X1 - 150
    LigneImageFleche1.Y1 = LigneImage.Y2
    LigneImageFleche2.X1 = LigneImage.X1
    LigneImageFleche2.X2 = LigneImage.X1 + 150
    LigneImageFleche2.Y1 = LigneImage.Y2

    If vI > 0 Then
        LigneImageFleche1.Y2 = LigneImage.Y2 + 150
        LigneImageFleche2.Y2 = LigneImage.Y2 + 150
    Else
        LigneImageFleche1.Y2 = LigneImage.Y2 - 150
        LigneImageFleche2.Y2 = LigneImage.Y2 - 150
    End If
    If LigneImage.Y2 > LigneImage.Y1 Then
        vPlus = 400
    Else
        vPlus = -400
    End If
    LblNomImage.Left = LigneImage.X1
    LblNomImage.Top = LigneImage.Y2 + vPlus
   ' LblNomImage.Visible = True

    '**************** Rayons pour convergente ****************'
    If OptType(0).Value = True Then
        LigneRayon(0).X1 = LigneObjet.X1 - 5000
        LigneRayon(0).X2 = LigneLentille.X1
        LigneRayon(0).Y1 = LigneObjet.Y2
        LigneRayon(0).Y2 = LigneObjet.Y2
    
        LigneRayon(1).X1 = LigneRayon(0).X2
        LigneRayon(1).X2 = LigneImage.X1 + 5000
        LigneRayon(1).Y1 = LigneRayon(0).Y2
        LigneRayon(1).Y2 = LigneImage.Y2 + 5000
    
        LigneRayon(2).X1 = LigneObjet.X1 - 10 * Abs(LigneObjet.X1 - LigneImage.X1)
        LigneRayon(2).X2 = LigneImage.X1
        LigneRayon(2).Y1 = LigneObjet.Y2
        LigneRayon(2).Y2 = LigneImage.Y2
    
        LigneRayon(3).X1 = LigneObjet.X1
        LigneRayon(3).X2 = LigneLentille.X1
        LigneRayon(3).Y1 = LigneObjet.Y2
        LigneRayon(3).Y2 = LigneImage.Y2
        LigneRayon(3).Visible = True
    
        LigneRayon(4).X1 = LigneRayon(3).X2
        LigneRayon(4).X2 = LigneImage.X1 + 5000
        LigneRayon(4).Y1 = LigneImage.Y2
        LigneRayon(4).Y2 = LigneImage.Y2
        LigneRayon(4).Visible = True

    Else
        '**************** Rayons pour divergente ****************'
        LigneRayon(0).X1 = LigneObjet.X1
        LigneRayon(0).X2 = LigneLentille.X1
        LigneRayon(0).Y1 = LigneObjet.Y2
        LigneRayon(0).Y2 = LigneObjet.Y2
    
        LigneRayon(1).X1 = LigneRayon(0).X2
        LigneRayon(1).X2 = LigneF.X1
        LigneRayon(1).Y1 = LigneRayon(0).Y2
        LigneRayon(1).Y2 = LigneFoyer.Y1
    
        LigneRayon(2).X1 = LigneObjet.X1
        LigneRayon(2).X2 = LigneLentille.X1
        LigneRayon(2).Y1 = LigneObjet.Y2
        LigneRayon(2).Y2 = LigneFoyer.Y1
    
        LigneRayon(3).Visible = False
        LigneRayon(4).Visible = False
    End If
End Function

Private Sub DessinFond()
    '**************** Lentille ****************'
    LigneLentille.X1 = FrmMain.Width / 2
    LigneLentille.X2 = FrmMain.Width / 2
    LigneLentille.Y1 = FrmMain.Height / 10
    LigneLentille.Y2 = FrmMain.Height / 10 * 8

    Dim vSigne As Integer
    If OptType(0).Value = True Then
        vSigne = -1
    Else
        vSigne = 1
    End If

    LigneType(0).X1 = LigneLentille.X1
    LigneType(0).Y1 = LigneLentille.Y1
    LigneType(0).X2 = LigneLentille.X1 - 200
    LigneType(0).Y2 = LigneLentille.Y1 - 200 * vSigne

    LigneType(1).X1 = LigneLentille.X1
    LigneType(1).Y1 = LigneLentille.Y1
    LigneType(1).X2 = LigneLentille.X1 + 200
    LigneType(1).Y2 = LigneLentille.Y1 - 200 * vSigne

    LigneType(2).X1 = LigneLentille.X1
    LigneType(2).Y1 = LigneLentille.Y2
    LigneType(2).X2 = LigneLentille.X1 - 200
    LigneType(2).Y2 = LigneLentille.Y2 + 200 * vSigne

    LigneType(3).X1 = LigneLentille.X1
    LigneType(3).Y1 = LigneLentille.Y2
    LigneType(3).X2 = LigneLentille.X1 + 200
    LigneType(3).Y2 = LigneLentille.Y2 + 200 * vSigne

    '**************** foyer ****************'
    LigneFoyer.X1 = 0
    LigneFoyer.X2 = FrmMain.Width
    LigneFoyer.Y1 = FrmMain.Height / 2
    LigneFoyer.Y2 = FrmMain.Height / 2

    '**************** Interface ****************'
    ChkInterface.Top = 200

    FrameType.Top = ChkInterface.Top + ChkInterface.Height + 100
    FrameType.Left = FrmMain.Width - FrameType.Width - 500

    ChkInterface.Left = FrameType.Left

    FrameValeurs.Top = FrameType.Top + FrameType.Height + 100
    FrameValeurs.Left = FrameType.Left
    
    FrameInfo.Top = FrameValeurs.Top + FrameValeurs.Height + 100
    FrameInfo.Left = FrameType.Left

    CmdDessiner.Top = FrameInfo.Height + FrameInfo.Top + 200
    CmdDessiner.Left = FrameType.Left

    CmdQuitter.Top = FrmMain.Height - 1000
    CmdQuitter.Left = FrameType.Left

    If ChkInterface.Value = Checked Then
        FrameType.Visible = True
        FrameValeurs.Visible = True
        FrameInfo.Visible = True
        CmdDessiner.Visible = True
        CmdQuitter.Visible = True
    Else
        FrameType.Visible = False
        FrameValeurs.Visible = False
        FrameInfo.Visible = False
        CmdDessiner.Visible = False
        CmdQuitter.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    DessinFond
    If Valeur = True Then
        If Calcul <> False Then
            DessinObjet
        End If
    End If
End Sub

Private Function Valeur() As Boolean
    If IsNumeric(TxtValF.Text) = False Or IsNumeric(TxtValP.Text) = False Or IsNumeric(TxtValO.Text) = False Then
        Valeur = False
    Else
        Valeur = True
    End If
End Function

Private Sub OptType_Click(Index As Integer)
DessinFond
    If Valeur = True Then
        If Calcul <> False Then
            DessinObjet
        End If
    End If
End Sub

Private Sub TxtValF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdDessiner_Click
    End If
End Sub

Private Sub TxtValp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdDessiner_Click
    End If
End Sub

Private Sub TxtValo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdDessiner_Click
    End If
End Sub

