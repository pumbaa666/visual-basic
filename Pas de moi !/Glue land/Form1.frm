VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dodo glu"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   1215
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   2143
      _Version        =   393216
      UpdateInterval  =   10
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      DeviceType      =   ""
      FileName        =   "C:\Maxi\Programation\visual basic\______glus land\ow2.wav"
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   480
      Top             =   3960
   End
   Begin VB.Frame Frame2 
      Caption         =   "Score :"
      Height          =   2055
      Left            =   0
      TabIndex        =   8
      Top             =   5040
      Width           =   2655
      Begin VB.Label Label3 
         Caption         =   "Votre temps :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bazouka                à pilules"
      Height          =   2055
      Left            =   2760
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Activer"
         Height          =   615
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton haut 
         Enabled         =   0   'False
         Height          =   615
         Left            =   840
         Picture         =   "Form1.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton bas 
         Enabled         =   0   'False
         Height          =   615
         Left            =   840
         Picture         =   "Form1.frx":16B4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton feu 
         Caption         =   "FEU !"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Droit 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1440
         Picture         =   "Form1.frx":1AF6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton gauche 
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         Picture         =   "Form1.frx":1F38
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   3240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bienvenue !!! Faites dormir les glus en leur donnant leurs pilules !!!"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   29
      Left            =   4440
      MouseIcon       =   "Form1.frx":237A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":35EC
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   28
      Left            =   0
      MouseIcon       =   "Form1.frx":3976
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4BE8
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   27
      Left            =   4200
      MouseIcon       =   "Form1.frx":4F72
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":61E4
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   26
      Left            =   2640
      MouseIcon       =   "Form1.frx":656E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":77E0
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   25
      Left            =   1920
      MouseIcon       =   "Form1.frx":7B6A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":8DDC
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   24
      Left            =   0
      MouseIcon       =   "Form1.frx":9166
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":A3D8
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   23
      Left            =   3000
      MouseIcon       =   "Form1.frx":A762
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":B9D4
      Top             =   4080
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   22
      Left            =   1800
      MouseIcon       =   "Form1.frx":BD5E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":CFD0
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   21
      Left            =   1560
      MouseIcon       =   "Form1.frx":D35A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":E5CC
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   20
      Left            =   1920
      MouseIcon       =   "Form1.frx":E956
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":FBC8
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   19
      Left            =   1320
      MouseIcon       =   "Form1.frx":FF52
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":111C4
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   18
      Left            =   840
      MouseIcon       =   "Form1.frx":1154E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":127C0
      Top             =   2640
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   17
      Left            =   3840
      MouseIcon       =   "Form1.frx":12B4A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":13DBC
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   16
      Left            =   120
      MouseIcon       =   "Form1.frx":14146
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":153B8
      Top             =   2640
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   15
      Left            =   480
      MouseIcon       =   "Form1.frx":15742
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":169B4
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   14
      Left            =   2400
      MouseIcon       =   "Form1.frx":16D3E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":17FB0
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   13
      Left            =   4440
      MouseIcon       =   "Form1.frx":1833A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":195AC
      Top             =   3360
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   12
      Left            =   2880
      MouseIcon       =   "Form1.frx":19936
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1ABA8
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   11
      Left            =   4080
      MouseIcon       =   "Form1.frx":1AF32
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1C1A4
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   10
      Left            =   960
      MouseIcon       =   "Form1.frx":1C52E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1D7A0
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   9
      Left            =   960
      MouseIcon       =   "Form1.frx":1DB2A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1ED9C
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   8
      Left            =   3840
      MouseIcon       =   "Form1.frx":1F126
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":20398
      Top             =   480
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   7
      Left            =   3480
      MouseIcon       =   "Form1.frx":20722
      Picture         =   "Form1.frx":21994
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   6
      Left            =   2520
      MouseIcon       =   "Form1.frx":21D1E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":22F90
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   5
      Left            =   0
      MouseIcon       =   "Form1.frx":2331A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":2458C
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   1680
      MouseIcon       =   "Form1.frx":24916
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":25B88
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   2880
      MouseIcon       =   "Form1.frx":25F12
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":27184
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   0
      MouseIcon       =   "Form1.frx":2750E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":28780
      Top             =   720
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   3360
      MouseIcon       =   "Form1.frx":28B0A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":29D7C
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   3240
      MouseIcon       =   "Form1.frx":2A106
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":2B378
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   1680
      Picture         =   "Form1.frx":2B702
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type glus
vitesse_x As Variant
vitesse_y As Variant
mort As Boolean
End Type
Private Type score
score As Long
nom As Variant
End Type
Private scores(5) As score
Private glu(30) As glus
Private deja(30) As Boolean
Private temps_départ
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub bas_Click()
Image2.top = Image2.top + 100
End Sub

Private Sub Command1_Click()
Image2.Visible = True
Command1.Enabled = False
feu.Enabled = True
Droit.Enabled = True
gauche.Enabled = True
haut.Enabled = True
bas.Enabled = True

End Sub

Private Sub Droit_Click()
Image2.left = Image2.left + 100
End Sub

Private Sub feu_Click()
Image2.Visible = False
Command1.Enabled = False
feu.Enabled = False
Droit.Enabled = False
gauche.Enabled = False
haut.Enabled = False
bas.Enabled = False


Dim dest As RECT
Dim r1 As RECT
Dim r2 As RECT
r1.top = Image2.top
r1.left = Image2.left
r1.Right = Image2.Width + Image2.left
r1.Bottom = Image2.Height + Image2.top
For b = 0 To 29
r2.top = Image1(b).top
r2.left = Image1(b).left
r2.Right = Image1(b).Width + Image1(b).left
r2.Bottom = Image1(b).Height + Image1(b).top

If IntersectRect(dest, r1, r2) Then
MMControl1.Command = "close"
MMControl1.FileName = App.Path & "\ow2.wav"
MMControl1.Command = "open"
MMControl1.Command = "Play"
glu(b).mort = True
Image1(b).Picture = LoadPicture(App.Path + "\mort.bmp")
Image1(b).MousePointer = 0
ggg = 0
For a = 0 To 30
If glu(a).mort Then ggg = ggg + 1
Next a
Label1 = ggg & " glu(s) qui dorment(dort)."
If ggg = 30 Then fin

End If
Next b

End Sub

Private Sub Form_Load()
MMControl1.FileName = App.Path & "\ow2.wav"
MMControl1.Command = "open"
scores(0).nom = GetSetting("glu", "scores", "nom1", "Personne")
scores(1).nom = GetSetting("glu", "scores", "nom2", "Personne")
scores(2).nom = GetSetting("glu", "scores", "nom3", "Personne")
scores(3).nom = GetSetting("glu", "scores", "nom4", "Personne")
scores(4).nom = GetSetting("glu", "scores", "nom5", "Personne")
scores(0).score = GetSetting("glu", "scores", "score1", "1000")
scores(1).score = GetSetting("glu", "scores", "score2", "1000")
scores(2).score = GetSetting("glu", "scores", "score3", "1000")
scores(3).score = GetSetting("glu", "scores", "score4", "1000")
scores(4).score = GetSetting("glu", "scores", "score5", "1000")

For a = 0 To 4
Label2 = Label2 & a + 1 & " : " & scores(a).nom & " --- " & scores(a).score & vbCrLf
Next a


Randomize Timer
For a = 0 To 29
asd = Int(Rnd * 50) + 1
glu(a).vitesse_x = asd - 10
asd = Int(Rnd * 50) + 1
glu(a).vitesse_y = asd - 10



Next a
milis = GetTickCount
secondes = milis \ 1000
milis = milis Mod 1000
minutes = secondes \ 60

temps_départ = secondes
End Sub



Private Sub gauche_Click()
Image2.left = Image2.left - 100
End Sub

Private Sub haut_Click()
Image2.top = Image2.top - 100
End Sub

Private Sub Image1_Click(Index As Integer)
glu(Index).mort = True
Image1(Index).Picture = LoadPicture(App.Path + "\mort.bmp")
Image1(Index).MousePointer = 0
MMControl1.Command = "close"
MMControl1.FileName = App.Path & "\ow2.wav"
MMControl1.Command = "open"
MMControl1.Command = "Play"

ggg = 0
For a = 0 To 30
If glu(a).mort Then ggg = ggg + 1
Next a
Label1 = ggg & " glu(s) qui dorment(dort)."
If ggg = 30 Then fin
End Sub



Private Sub Timer1_Timer()

For a = 0 To 29
asdf = False
If Image1(a).left < 0 Then glu(a).vitesse_y = -glu(a).vitesse_y: asdf = True
If Image1(a).top < 0 Then glu(a).vitesse_x = -glu(a).vitesse_x: asdf = True
If Image1(a).left > 4800 Then glu(a).vitesse_y = -glu(a).vitesse_y: asdf = True
If Image1(a).top > 4320 Then glu(a).vitesse_x = -glu(a).vitesse_x: asdf = True
deja(a) = asdf
Dim dest As RECT
Dim r1 As RECT
Dim r2 As RECT
r1.top = Image1(a).top
r1.left = Image1(a).left
r1.Right = Image1(a).Width + Image1(a).left
r1.Bottom = Image1(a).Height + Image1(a).top

For b = 0 To 29
If b = a Then GoTo sui
r2.top = Image1(b).top
r2.left = Image1(b).left
r2.Right = Image1(b).Width + Image1(b).left
r2.Bottom = Image1(b).Height + Image1(b).top

If IntersectRect(dest, r1, r2) Then
If deja(a) Then
Else
glu(a).vitesse_x = -glu(a).vitesse_x
glu(a).vitesse_y = -glu(a).vitesse_y
deja(a) = True
End If
'glu(b).vitesse_x = -glu(b).vitesse_x
'glu(b).vitesse_y = -glu(b).vitesse_y
'MsgBox a, , b
End If

sui:
Next b

Next a

For a = 0 To 29
If Not (glu(a).mort) Then
Image1(a).left = Image1(a).left + glu(a).vitesse_y
Image1(a).top = Image1(a).top + glu(a).vitesse_x
deja(a) = False
End If
Next a

End Sub

Private Sub Timer2_Timer()

milis = GetTickCount
secondes = milis \ 1000
milis = milis Mod 1000
minutes = secondes \ 60

secondes = secondes - temps_départ
Label3 = "Temps actuel : " & secondes
End Sub
Public Sub fin()
Dim secondes As Long
milis = GetTickCount
secondes = milis \ 1000
milis = milis Mod 1000
minutes = secondes \ 60

secondes = secondes - temps_départ

For a = 0 To 4
If secondes < scores(a).score Then
nom = InputBox("Vous êtes dans les meileurs scores !!! (" & a & ")", "Entez votre nom :", "ICI !   ;-)")
For b = 5 To a Step -1
If b <> 0 Then scores(b) = scores(b - 1)
Next b
scores(a).score = secondes
scores(a).nom = nom

 SaveSetting "glu", "scores", "nom1", scores(0).nom
 SaveSetting "glu", "scores", "nom2", scores(1).nom
 SaveSetting "glu", "scores", "nom3", scores(2).nom
 SaveSetting "glu", "scores", "nom4", scores(3).nom
 SaveSetting "glu", "scores", "nom5", scores(4).nom
 SaveSetting "glu", "scores", "score1", scores(0).score
SaveSetting "glu", "scores", "score2", scores(1).score
 SaveSetting "glu", "scores", "score3", scores(2).score
 SaveSetting "glu", "scores", "score4", scores(3).score
 SaveSetting "glu", "scores", "score5", scores(4).score
GoTo fin
End If

Next a
fin:
End


End Sub

