VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Testez votre vitesse de frappe"
   ClientHeight    =   4395
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dial 
      Left            =   6480
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt"
   End
   Begin VB.CommandButton CmdTexte 
      Caption         =   "&Choisir un texte"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox TxtOriginal 
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Vitesse.frx":0000
   End
   Begin VB.CommandButton CmdCommencer 
      Caption         =   "&Commencer"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Timer ClkTemps 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   1200
   End
   Begin VB.ComboBox ComboDuree 
      Height          =   315
      ItemData        =   "Vitesse.frx":0082
      Left            =   6000
      List            =   "Vitesse.frx":008F
      TabIndex        =   2
      Text            =   "Durée de l'épreuve"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox TxtNew 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label Label4 
      Caption         =   "sec"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.Label LblSec 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "min"
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label LblMin 
      Caption         =   "0"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Temps restant :"
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu MenuFichierOuvrir 
         Caption         =   "Ouvrir"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Tiret1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu menuTemps 
      Caption         =   "Temps"
      Begin VB.Menu MenuTemps1 
         Caption         =   "1 min"
      End
      Begin VB.Menu MenuTemps5 
         Caption         =   "5 min"
      End
      Begin VB.Menu MenuTemps10 
         Caption         =   "10 min"
      End
   End
   Begin VB.Menu MenuAide 
      Caption         =   "Aide"
      Begin VB.Menu MenuAideAide 
         Caption         =   "Aide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuAideAbout 
         Caption         =   "A propos"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCount As Integer
Dim vNbFautes As Integer

Private Sub ClkTemps_Timer()
Static vCntTemps As Integer
    LblSec.Caption = Int(LblSec.Caption) - 1
    If Int(LblSec.Caption) = -1 Then
        LblSec.Caption = "59"
        LblMin.Caption = Int(LblMin.Caption) - 1
        If Int(LblMin.Caption) = -1 Then
            LblMin.Caption = "0"
            LblSec.Caption = "0"
            ClkTemps.Enabled = False
            MsgBox "Terminé"
            ComboDuree.Enabled = True
            CmdCommencer.Caption = "&Commencer"

            FrmStatistiques.LblNbFrappes = "Nombre de frappes : " & Len(TxtNew.Text)
            FrmStatistiques.LblBrute = "Vitesse brute : " & Left(Len(TxtNew.Text) / (60 * Val(ComboDuree.Text)), 3) & " frappes par secondes"
            FrmStatistiques.LblFautes = "Nombre de fautes : " & vNbFautes
            FrmStatistiques.LblReel = "Vitesse pure : " & Left((Len(TxtNew.Text) - 10 * vNbFautes) / (60 * Val(ComboDuree.Text)), 3) & " frappes par secondes"
            FrmStatistiques.Show
        End If
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdCommencer_Click()
    If CmdCommencer.Caption = "&Pause [Esc]" Then
        ClkTemps.Enabled = False
        CmdCommencer.Caption = "&Reprendre"
    ElseIf CmdCommencer.Caption = "&Reprendre" Then
        ClkTemps.Enabled = True
        CmdCommencer.Caption = "&Pause [Esc]"
        TxtNew.SetFocus
    Else
        If TxtOriginal.Text = "" Then
            MsgBox "Choisissez un text", vbCritical, "Erreur"
        ElseIf ComboDuree.Text = "Durée de l'épreuve" Then
            MsgBox "Choisissez la durée de l'épreuve", vbCritical, "Erreur"
        Else
            TxtOriginal.Font = RGB(0, 0, 0)
            TxtOriginal.SelStart = 1
            TxtOriginal.SelLength = 1
            TxtNew.Text = ""
            vCount = 1
            TxtNew.Enabled = True
            LblMin.Caption = Val(ComboDuree)
            LblSec.Caption = "0"
            ComboDuree.Enabled = False
            CmdCommencer.Caption = "&Pause [Esc]"
            TxtNew.SetFocus
        End If
    End If
End Sub

Private Sub CmdTexte_Click()
Dim vLigne As String

    Dial.ShowOpen
    If Dial.FileName <> "" Then
        If Right(Dial.FileName, 4) <> ".txt" Then
            MsgBox "Format inconnu", vbCritical, "Erreur*"
        Else
            TxtOriginal.Text = ""
            Open Dial.FileName For Input As #1
            Do
                Line Input #1, vLigne
                TxtOriginal.Text = TxtOriginal.Text & vLigne & Chr(13)
            Loop While Not (EOF(1))
            Close #1
        End If
    End If
End Sub


Private Sub Form_Load()
    TxtOriginal.Font = RGB(0, 0, 0)
End Sub

Private Sub MenuAideAbout_Click()
    FrmAbout.Show
    FrmMain.Hide
End Sub

Private Sub MenuAideAide_Click()
    FrmAide.Show
    FrmMain.Hide
End Sub

Private Sub MenuFichierOuvrir_Click()
    CmdTexte_Click
End Sub

Private Sub MenuFichierQuitter_Click()
    End
End Sub

Private Sub MenuTemps1_Click()
    MenuTemps1.Checked = True
    MenuTemps5.Checked = False
    MenuTemps10.Checked = False

    ComboDuree.Text = "1 min"
End Sub

Private Sub MenuTemps5_Click()
    MenuTemps1.Checked = False
    MenuTemps5.Checked = True
    MenuTemps10.Checked = False

    ComboDuree.Text = "5 min"
End Sub

Private Sub MenuTemps10_Click()
    MenuTemps1.Checked = False
    MenuTemps5.Checked = False
    MenuTemps10.Checked = True

    ComboDuree.Text = "10 min"
End Sub

Private Sub TxtNew_KeyPress(KeyAscii As Integer)
Static vLastFautes As Integer

    If KeyAscii = 27 And CmdCommencer.Caption = "&Pause [Esc]" Then
        CmdCommencer_Click
    ElseIf KeyAscii = 32 Or KeyAscii = 13 Then
        LireMot (vCount)
    End If
    If TxtNew.Text = "" Then
        ClkTemps.Enabled = True
        LireMot (0)
        If Chr(KeyAscii) <> Mid(TxtOriginal.Text, vCount, 1) Then
            KeyAscii = 0
            If vCount <> vLastFautes Then
                vNbFautes = vNbFautes + 1
                vLastFautes = vCount
            End If
        Else
            vCount = vCount + 1
        End If
    Else
        If Chr(KeyAscii) <> Mid(TxtOriginal.Text, vCount, 1) Then
            KeyAscii = 0
            If vCount <> vLastFautes Then
                vNbFautes = vNbFautes + 1
                vLastFautes = vCount
            End If
        Else
            vCount = vCount + 1
        End If
    End If
End Sub

Function LireMot(ByVal vNum As Integer)
Dim vMot As String
Dim vStart As Integer
Dim vNbFois As Integer
Dim vFin As Integer
    vFin = vNum + 20
    vStart = vNum
    Do
        If Mid(TxtOriginal.Text, vNum + 1, 1) = " " Or Mid(TxtOriginal.Text, vNum + 1, 1) = Chr(13) Then
            TxtOriginal.Font = RGB(0, 0, 0)
            TxtOriginal.SelStart = vStart
            TxtOriginal.SelLength = vFin - vStart
            If vNbFois = 1 Then
                TxtOriginal.SelColor = RGB(255, 0, 0)
                vNum = 0
            Else
                vNbFois = 1
                vFin = vNum
                vNum = vNum + 1
            End If
        Else
            vMot = vMot & Mid(TxtOriginal.Text, vNum + 1, 1)
            vNum = vNum + 1
        End If
    Loop While (vNum <> 0)
End Function
