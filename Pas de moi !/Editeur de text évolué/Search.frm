VERSION 5.00
Begin VB.Form Search 
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox CheckReplace 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8D0C0&
      Caption         =   "Remplacer par :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox TextSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox TextReplace 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin PN3.Button ButtonCancel 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   192
      BackColor       =   14938367
      Caption         =   "Annuler"
      ForeColor       =   192
   End
   Begin PN3.Button ButtonSearch 
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   6050618
      BackColor       =   15526369
      Caption         =   "Rechercher"
   End
   Begin VB.Label LabelTitle 
      BackStyle       =   0  'Transparent
      Caption         =   " Recherche et remplacement"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   10
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechercher :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00D8D0C0&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00B4A587&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LastPlace As Single
Dim LastRecherche As String

Private Sub ButtonCancel_Click()

Me.Hide
Main.Enabled = True
Main.Text.SetFocus
TextSearch.Text = ""
TextReplace.Text = ""
TextReplace.Enabled = False
CheckReplace.Value = 0
CheckReplace.ForeColor = &H80000011
TextReplace.BackColor = &HE0E0E0
    
End Sub

Private Sub ButtonSearch_Click()

If TextSearch.Text = "" Then MessageBox.Message "Aucun texte à chercher.", "Erreur durant la recherche", OkOnly, Information, Search: Exit Sub

ButtonSearch.Enabled = False

If CheckReplace.Value = 0 Then
    If SearchText(Main.Text.Text, TextSearch.Text) = False Then Me.Hide: MessageBox.Message "La recherche n'a rien donné.", "Recherche infructueuse", OkOnly, Information, Search
Else
    ReplaceText Main.Text.Text, TextSearch.Text, TextReplace.Text
End If

ButtonSearch.Enabled = True
LastRecherche = TextSearch.Text

ButtonCancel_Click

End Sub

Private Sub CheckReplace_Click()

If CheckReplace.Value = 0 Then
    CheckReplace.ForeColor = &H80000011
    TextReplace.BackColor = &HE0E0E0
    TextReplace.Enabled = False
Else
    CheckReplace.ForeColor = &H80000008
    TextReplace.BackColor = &H80000005
    TextReplace.Enabled = True
End If

End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour bouger la fenêtre
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Function SearchText(Texte As String, Recherche As String) As Boolean

Dim Incr As Single

'Fonction de recherche de texte

For Incr = 1 To Len(Texte)
    If LCase$(Recherche) = LCase$(Mid$(Texte, Incr, Len(Recherche))) Then
        Main.Text.SelStart = Incr - 1
        Main.Text.SelLength = Len(Recherche)
        SearchText = True
        LastPlace = Incr + Len(Recherche)
        Exit Function
    End If
Next

End Function

Public Function SearchNextText() As Boolean

Dim Incr As Single
Dim Texte As String

'Fonction de recherche de texte par appui sur F3

Texte = Main.Text.Text

For Incr = LastPlace To Len(Texte)
    If LCase$(LastRecherche) = LCase$(Mid$(Texte, Incr, Len(LastRecherche))) Then
        Main.Text.SelStart = Incr - 1
        Main.Text.SelLength = Len(LastRecherche)
        SearchNextText = True
        LastPlace = Incr + Len(LastRecherche)
        Exit Function
    End If
Next

End Function

Private Sub ReplaceText(Texte As String, Recherche As String, Remplacement As String)

Dim Incr As Single

'Fonction de remplacement de texte

For Incr = 1 To Len(Texte)
    If LCase(Recherche) = LCase(Mid$(Texte, Incr, Len(Recherche))) Then
        Main.Text.SelStart = Incr - 1
        Main.Text.SelLength = Len(Recherche)
        Main.Text.SelText = Remplacement
    End If
Next

End Sub

Private Sub TextSearch_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then ButtonSearch_Click

End Sub

Private Sub TextReplace_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then ButtonSearch_Click

End Sub
