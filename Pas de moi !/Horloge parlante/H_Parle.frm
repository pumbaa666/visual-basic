VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Horloge parlante  ryl..."
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00DDFFDD&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00DDFFDD&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   7455
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   200
      Top             =   2400
   End
   Begin MCI.MMControl MMControl1 
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3480
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1085
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Donner l'heure"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Line Line4 
      X1              =   7800
      X2              =   7800
      Y1              =   120
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   7800
      X2              =   120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   2520
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7800
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Horloge parlante pour votre voix ,voir répertoire "son_horl"
'rylou , 23:08 06/05/2004

Dim Chemin As String        'Chemin du répertoire des fichiers "*.WAV"
Dim T As Integer

Private Sub Command1_Click()
Chemin = App.Path + "\Son_Horl\"

T = Hour(Time)                      'HEURE dans T
Call Corresp                    'TROUVER LE SON DU NOMBRE
MMControl1.FileName = Chemin + "H.wav" 'LECTURE (Heure)
Call Lecture

T = Minute(Time)                    'MINUTE dans T
Call Corresp                    'TROUVER LE SON DU NOMBRE
MMControl1.FileName = Chemin + "M.wav" 'LECTURE (Minute)
Call Lecture

T = Second(Time)                    'SECONDE dans T
Call Corresp                    'TROUVER LE SON DU NOMBRE
MMControl1.FileName = Chemin + "S.wav" 'LECTURE (Seconde)
Call Lecture
End Sub
Sub Corresp()
If T = 1 Then MMControl1.FileName = Chemin + "une.wav"
If T <= 20 And T <> 1 Then MMControl1.FileName = Chemin + Left$(T, 2) + ".wav"

If T <= 29 And T >= 21 Then 'LECTURE NOMBRE equiv <21...29>
MMControl1.FileName = Chemin + "20.wav"
Call Lecture
MMControl1.FileName = Chemin + Right$(T, 1) + ".wav"
End If
If T <= 39 And T >= 30 Then 'LECTURE NOMBRE equiv <30...39>
MMControl1.FileName = Chemin + "30.wav"
If T <> 30 Then Call Lecture
If T <> 30 Then MMControl1.FileName = Chemin + Right$(T, 1) + ".wav"
End If
If T <= 49 And T >= 40 Then 'LECTURE NOMBRE equiv <40...49>
MMControl1.FileName = Chemin + "40.wav"
If T <> 40 Then Call Lecture
If T <> 40 Then MMControl1.FileName = Chemin + Right$(T, 1) + ".wav"
End If
If T <= 59 And T >= 50 Then 'LECTURE NOMBRE equiv <50...59>
MMControl1.FileName = Chemin + "50.wav"
If T <> 50 Then Call Lecture
If T <> 50 Then MMControl1.FileName = Chemin + Right$(T, 1) + ".wav"
End If
Call Lecture 'LECTURE NOMBRE equiv de 0...20 ou l'unitee de T si>21
End Sub

Sub Lecture()
MMControl1.Command = "open"
MMControl1.Wait = True          'ATTENDRE la fin de la lecture
MMControl1.Command = "play"     'LECTURE
MMControl1.Command = "close"    'FERMETURE
Timer1_Timer                    'RAFRAICHIR TEXT1.TEXT
End Sub

Private Sub Timer1_Timer()
Text1.Text = Time
Text2.Text = Date
Text1.Refresh
End Sub
