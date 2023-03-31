VERSION 5.00
Begin VB.Form Crypt 
   BorderStyle     =   0  'None
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin PN3.Progress Progress 
      Height          =   255
      Left            =   120
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Appearance      =   0
      Value           =   0
      ColorBar        =   15526369
      BackColor       =   14209216
   End
   Begin PN3.Button ButtonCancel 
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   192
      BackColor       =   14938367
      Caption         =   "Annuler"
      ForeColor       =   192
   End
   Begin PN3.Button ButtonDCrypt 
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   6050618
      BackColor       =   15526369
      Caption         =   "Décrypter"
   End
   Begin PN3.Button ButtonCrypt 
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   6050618
      BackColor       =   15526369
      Caption         =   "Crypter"
   End
   Begin VB.TextBox TextBoucles 
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
      Left            =   2040
      TabIndex        =   6
      Text            =   "1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TextPass2 
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
      Left            =   1320
      PasswordChar    =   "•"
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox TextPass1 
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
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de boucles :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Retapez-le :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label LabelTitle 
      BackStyle       =   0  'Transparent
      Caption         =   " Cryptage/Décryptage"
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
      TabIndex        =   0
      Top             =   15
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
   Begin VB.Shape Shape 
      BackColor       =   &H00D8D0C0&
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Crypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Annulé As Boolean

Private Sub ButtonCancel_Click()

Annulé = True
Me.Hide
Main.Enabled = True
Main.Text.SetFocus
TextPass1.Text = ""
TextPass2.Text = ""
TextBoucles.Text = "1"

End Sub

Private Sub ButtonCrypt_Click()

Dim IncrX As Single

If VerifIsGood = False Then Exit Sub
Annulé = False
LockCrypt

For IncrX = 1 To Val(TextBoucles.Text)
    Encrypt Main.Text.Text, TextPass1.Text
Next

MustBeSaved = True
UnlockCrypt
ButtonCancel_Click

End Sub

Private Sub ButtonDCrypt_Click()

Dim IncrX As Single

If VerifIsGood = False Then Exit Sub
Annulé = False
LockCrypt

For IncrX = 1 To Val(TextBoucles.Text)
    Decrypt Main.Text.Text, TextPass1.Text
Next

MustBeSaved = True
UnlockCrypt
ButtonCancel_Click

End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour bouger la fenêtre
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Function Encrypt(Text As String, Pass As String) As String

Dim A, B, C As Integer
Dim Buffer As String
Dim Incr As Single

'Fonction de cryptage

Progress.Maxi = Len(Text)

For Incr = 1 To Len(Text)
    DoEvents
    If Annulé = True Then Annulé = False: Progress.Value = 0: Exit Function
    A = A + 1
    If A > Len(Pass) Then A = 1
    B = (Asc(Mid$(Text, Incr, 1))) + (Asc(Mid$(Pass, A, 1)))
    If B > 255 Then B = B - 255
    Buffer = Buffer & Chr(B)
    C = C + 1
    If C = 500 Then Progress.Value = Progress.Value + C: Encrypt = Encrypt & Buffer: Buffer = "": C = 0
Next

If C <> 0 Then
    Progress.Value = Progress.Value + C
    Encrypt = Encrypt & Buffer
End If

Main.Text.Text = Encrypt
Progress.Value = 0

End Function

Private Function Decrypt(Text As String, Pass As String) As String

Dim A, B, C As Integer
Dim Buffer As String
Dim Incr As Single

'Fonction de décryptage

Progress.Maxi = Len(Text)

For Incr = 1 To Len(Text)
    DoEvents
    If Annulé = True Then Annulé = False: Progress.Value = 0: Exit Function
    A = A + 1
    If A > Len(Pass) Then A = 1
    B = (Asc(Mid$(Text, Incr, 1))) - (Asc(Mid$(Pass, A, 1)))
    If B < 0 Then B = B + 255
    Buffer = Buffer & Chr(B)
    C = C + 1
    If C = 500 Then Progress.Value = Progress.Value + C: Decrypt = Decrypt & Buffer: Buffer = "": C = 0
Next

If C <> 0 Then
    Progress.Value = Progress.Value + C
    Decrypt = Decrypt & Buffer
End If

Main.Text.Text = Decrypt
Progress.Value = 0

End Function

Private Function VerifIsGood() As Boolean

'Fonction de vérification des champs
If TextPass1.Text = "" Or TextPass2.Text = "" Or TextBoucles.Text = "" Then MessageBox.Message "Remplissez tous les champs.", "Erreur dans les champs", OkOnly, Information, Crypt: Exit Function
If TextPass1.Text <> TextPass2.Text Then MessageBox.Message "Les mots de passe ne correspondent pas." & vbCrLf & "Veuillez retapez votre mot de passe.", "Mot de passe invalide", OkOnly, Information, Crypt: TextPass1.Text = "": TextPass2.Text = "": Exit Function
If Val(TextBoucles.Text) < 1 Then MessageBox.Message "Le nombre de boucles ne peut être inférieur à 1.", "Erreur", OkOnly, Information, Crypt: TextBoucles.Text = "": Exit Function

VerifIsGood = True

End Function

Private Sub LockCrypt()

ButtonCrypt.Enabled = False
ButtonDCrypt.Enabled = False
TextPass1.Enabled = False
TextPass2.Enabled = False
TextBoucles.Enabled = False
TextPass1.BackColor = &HE0E0E0
TextPass2.BackColor = &HE0E0E0
TextBoucles.BackColor = &HE0E0E0

End Sub

Private Sub UnlockCrypt()

ButtonCrypt.Enabled = True
ButtonDCrypt.Enabled = True
TextPass1.Enabled = True
TextPass2.Enabled = True
TextBoucles.Enabled = True
TextPass1.BackColor = &H80000005
TextPass2.BackColor = &H80000005
TextBoucles.BackColor = &H80000005

End Sub

