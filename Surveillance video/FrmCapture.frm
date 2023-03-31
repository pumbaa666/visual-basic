VERSION 5.00
Begin VB.Form FrmCapture 
   Caption         =   "Capture d'image"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNow 
      Caption         =   "&Maintenant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Cliquez ici pour prendre une photo maintenant"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox TxtDegre 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "180"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox TxtDegre 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "90"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox TxtDegre 
      Alignment       =   2  'Center
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
      Index           =   0
      Left            =   360
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Capturer une image quand la caméra arrive aux positions: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "FrmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNbVirgule As Boolean
Dim tVal(2) As Integer

Private Sub CmdAnnuler_Click()
Dim vCntVal As Integer
' Remet les valeurs initiales dans les TextBox
    For vCntVal = 0 To 2
        TxtDegre(vCntVal).Text = tVal(vCntVal)
    Next
    FrmMain.Show
    FrmCapture.Hide
End Sub

Private Sub CmdNow_Click()
Dim vTemp As Boolean
Dim vChemin As String
' Entre .avi à la fin du nom de fichier si ce n'est pas déjà fait
    vChemin = InputBox("Entrez le chemin et le nom du fichier", "Capture d'image")
    If Right(vChemin, 4) <> ".bmp" Then
        vChemin = vChemin & ".bmp"
    End If
    
' Enregistre l'image
    frmMainCam.VideoPortal1.PictureToFile 0, 24, vChemin, ""
End Sub

Private Sub CmdOk_Click()
Dim vCount As Integer
    vNbVirgule = False
    For vCount = 0 To 2
        If TxtDegre(vCount).Text = "" Then
            MsgBox "Il manque une ou plusieurs valeurs!!!", vbCritical, "Erreur"
            vCount = 2
        ElseIf Int(TxtDegre(vCount).Text) < 0 Or Int(TxtDegre(vCount).Text) > 180 Then
            MsgBox "L'angle doit être compris entre 0 et 180°!", vbCritical, "Erreur"
            vCount = 2
        ElseIf TxtDegre(0).Text = TxtDegre(1).Text Or TxtDegre(0).Text = TxtDegre(2).Text Or TxtDegre(1).Text = TxtDegre(2).Text Then
            MsgBox "Entrez des valeurs différentes pour chaque positions!", vbCritical, "Erreur"
            vCount = 2
        Else

' Sauvegarde les valeurs des TextBox dans un tableau
            tArret(vCount) = Int(TxtDegre(vCount).Text)
            FrmMain.Show
            FrmCapture.Hide
        End If
    Next
End Sub


Private Sub Form_Activate()
Dim vCntVal As Integer
' Sauvegarde les valeurs des TextBox dans un tableau pour les réstaurer si on click sur Annuler
    For vCntVal = 0 To 2
        tVal(vCntVal) = Int(TxtDegre(vCntVal).Text)
    Next
End Sub

Private Sub TxtDegre_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 And KeyAscii <> 46) Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur de saisie"
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        CmdOk_Click
    ElseIf KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
    If KeyAscii = 44 Then
        If vNbVirgule = False Then
            vNbVirgule = True
        Else
            KeyAscii = 0
        End If
        If Len(TxtDegre(Index).Text) = 0 Then
            vNbVirgule = False
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = 8 And Right(TxtDegre(Index).Text, 1) = "," Then
        vNbVirgule = False
    End If
End Sub
