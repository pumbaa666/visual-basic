VERSION 5.00
Begin VB.Form FrmPixel 
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   2400
      Left            =   3120
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Veuillez patienter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Blanc"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "255, 255, 255"
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Noir"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "0, 0, 0, "
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer ClkLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   240
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TxtPixel 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "B"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox TxtPixel 
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "G"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox TxtPixel 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "R"
      Top             =   2160
      Width           =   495
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Autre..."
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Violet"
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "128, 0, 128"
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Orange"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "255, 128, 64"
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Vert"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      ToolTipText     =   "0, 255, 0"
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Jaune"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "255, 255, 0"
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Bleu"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "0, 0, 255"
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton OptCouleur 
      Caption         =   "Rouge"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "255, 0, 0"
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Label LblFormat 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Choisissez la couleur à rechercher"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "FrmPixel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTemp As Byte
Dim vLong As Integer
Dim vFichier As String
Dim vNomCoul As String

Private Sub ClkLoad_Timer()
Dim vCount As Long
Dim vCount2 As Integer
Dim vCount3 As Integer
Dim vNbPix As Long ' 0 = Bleu   1 = Vert   2 = Jaune
Dim vCouleur As String
Dim vRecherche As String
vCount2 = 1
    
    If OptCouleur(0).Value = True Then
        vRecherche = "255, 0, 0, "
    ElseIf OptCouleur(1).Value = True Then
        vRecherche = "0, 0, 255, "
    ElseIf OptCouleur(2).Value = True Then
        vRecherche = "255, 255, 0, "
    ElseIf OptCouleur(3).Value = True Then
        vRecherche = "0, 255, 0, "
    ElseIf OptCouleur(4).Value = True Then
        vRecherche = "255, 128, 64, "
    ElseIf OptCouleur(5).Value = True Then
        vRecherche = "128, 0, 128, "
    ElseIf OptCouleur(6).Value = True Then
        vRecherche = TxtPixel(0) & ", " & TxtPixel(1) & ", " & TxtPixel(2) & ", "
    ElseIf OptCouleur(7).Value = True Then
        vRecherche = "0, 0, 0, "
    ElseIf OptCouleur(8).Value = True Then
        vRecherche = "255, 255, 255, "
    End If
      
    Open vFichier For Random As #1 Len = 1
    For vCount = 55 To FileLen(vFichier)
        If vCount = 55 + vLong * (3 + vCount3) Then
            vCount = vCount + 3
            vCount3 = vCount3 + 1
        End If
        Get #1, vCount, vTemp
        vCouleur = vTemp & ", " & vCouleur
        vCount2 = vCount2 + 1
        If vCount2 = 4 Then
            If vCouleur = vRecherche Then
                vNbPix = vNbPix + 1
            End If
            List.AddItem Left(vCouleur, Len(vCouleur) - 2)
            vCouleur = ""
            vCount2 = 1
        End If
    Next
    Close #1
    ClkLoad.Enabled = False
    CmdLoad.Visible = False
'    If OptCouleur(7).Value = True Then
'        vNbPix = vNbPix - vLong
'    End If
    MsgBox "Il y a " & vNbPix & " pixels " & vNomCoul & "s.", vbInformation, "Pixels"
    If MsgBox("Voulez vous voir le détails?!?", vbYesNo, "Détails") = vbYes Then
        FrmPixel.Width = 4700
        List.Visible = True
    End If
End Sub

Private Sub CmdAnnuler_Click()
    FrmExplorateur.Show
    FrmPixel.Hide
End Sub

Private Sub CmdOk_Click()
Dim vTest As Boolean

    List.Clear
    List.Visible = False
    FrmPixel.Width = 3300
    If TxtPixel(0).Enabled = True Then
        For vCount = 0 To 2
            If TxtPixel(vCount).Text <> "" And TxtPixel(0) <> "R" And TxtPixel(1) <> "G" And TxtPixel(2) <> "B" Then
                If Int(TxtPixel(vCount).Text) < 0 Or Int(TxtPixel(vCount).Text) > 255 Then
                    vTest = True
                End If
            Else
                vTest = True
            End If
        Next
    End If
    
    If vTest = True Then
        MsgBox "Chaque n° doit être compris entre 0 et 255", vbCritical, "Erreur"
    Else
        CmdLoad.Visible = True
        ClkLoad.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
Dim vHaut As Integer
    If Len(FrmExplorateur.Dir1.Path) = 3 Then
        vFichier = FrmExplorateur.Dir1.Path & FrmExplorateur.File1.FileName
    Else
        vFichier = FrmExplorateur.Dir1.Path & "\" & FrmExplorateur.File1.FileName
    End If
    
    Open vFichier For Random As #1 Len = 1
    Get #1, 19, vTemp
    vLong = vTemp
    Get #1, 20, vTemp
    vLong = vLong + 256 * vTemp
    
    Get #1, 23, vTemp
    vHaut = vTemp
    Get #1, 24, vTemp
    vHaut = vHaut + 256 * vTemp
    Close #1
    LblFormat.Caption = "Format de l'image : " & vLong & " X " & vHaut
    
    FrmPixel.Caption = vFichier
    vNomCoul = "Rouge"
    FrmPixel.Width = 3300
    List.Visible = False
End Sub

Private Sub List_Click()
    MsgBox List.ListIndex
End Sub

Private Sub OptCouleur_Click(Index As Integer)
    If OptCouleur(6).Value = True Then
        TxtPixel(0).Enabled = True
        TxtPixel(1).Enabled = True
        TxtPixel(2).Enabled = True
    Else
        TxtPixel(0).Enabled = False
        TxtPixel(1).Enabled = False
        TxtPixel(2).Enabled = False
    End If
    vNomCoul = OptCouleur(Index).Caption
    FrmPixel.Width = 3300
    List.Visible = False
End Sub

Private Sub TxtPixel_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        MsgBox "Veuillez n'entrer que des chiffres", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub
