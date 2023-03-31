VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Computerfont"
      Size            =   33
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   3360
   End
   Begin RichTextLib.RichTextBox R1 
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      _Version        =   393217
      BackColor       =   0
      Enabled         =   0   'False
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Computerfont"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Create By Jielde"
      BeginProperty Font 
         Name            =   "Computerfont"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCaption As String
Dim i As Long
Dim ColorR1 As String
Dim couleur
Dim r
Dim ValCaption As Long



Private Sub Form_KeyPress(KeyAscii As Integer)
'Si Echap alors quitte le programme.
If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()
'----------------------------------------------------------
'Ajuste la taille de la fenetre avec celle de l'écrant.
Me.Height = Screen.Height
Me.Width = Screen.Width
'----------------------------------------------------------
ValCaption = 0
load_form
End Sub

Function load_form()

mCaption = mcaption_set(ValCaption)
Me.ForeColor = vbBlack
Me.FontSize = "20"
Me.CurrentX = Me.ScaleWidth / 2 - TextWidth(mCaption) / 2
Me.CurrentY = Me.ScaleHeight / 2 - TextHeight(mCaption) / 2
Print mCaption

Label1.Caption = mCaption
R1.SelStart = 0: R1.SelLength = Len(R1.Text): R1.SelColor = vbBlack
R1.Refresh
R1.Width = Label1.Width + 100

R1.Left = (Me.ScaleWidth / 2) - (R1.Width / 2)
R1.Top = Me.ScaleHeight / 2 - (R1.Height / 2)
R1.Text = mCaption
i = 0
ColorR1 = vbWhite
End Function

Private Sub Timer1_Timer()
If i = Len(R1.Text) + 1 Then
Timer1.Enabled = False
r = 0
Timer2.Enabled = True
    i = 0
    If ColorR1 = vbBlack Then
        ColorR1 = vbWhite
    Else
        ColorR1 = vbBlack
    End If
    Exit Sub
End If
Timer1.Interval = 100
R1.SelStart = 0
R1.SelLength = i
R1.SelColor = vbBlack
R1.SelStart = i
R1.SelLength = 2
R1.SelColor = vbWhite
R1.SelLength = 0
i = i + 1
End Sub

Function degrad_couleur()
For r = 0 To 255
    couleur = RGB(r, r, r)
    R1.SelStart = 0: R1.SelLength = Len(R1.Text): R1.SelColor = couleur
    DoEvents
Next r
End Function

Private Sub Timer2_Timer()
If r = 255 Then
    Timer3.Interval = 1000
    Timer2.Enabled = False
    Timer3.Enabled = True
    r = 0
    Exit Sub
End If
    couleur = RGB(r, r, r)
    R1.SelStart = 0: R1.SelLength = Len(R1.Text): R1.SelColor = couleur
    DoEvents
    r = r + 1
End Sub

Private Sub Timer3_Timer()
If i = Len(R1.Text) + 1 Then
    Timer3.Enabled = False
    i = 0
    ValCaption = ValCaption + 1
    load_form
    Timer1.Enabled = True
    If ColorR1 = vbBlack Then
        ColorR1 = vbWhite
    Else
        ColorR1 = vbBlack
    End If
    Exit Sub
End If
Timer3.Interval = 100
R1.SelStart = 0
R1.SelLength = i
R1.SelColor = vbBlack
R1.SelStart = i
R1.SelLength = 2
R1.SelColor = vbWhite
R1.SelLength = 0
i = i + 1
End Sub

Function mcaption_set(captionVal As Long)
Select Case captionVal
    Case 0
        mcaption_set = "Jielde Present"
    
    Case 1
        mcaption_set = "Mission:Intro"
    
    Case 2
        mcaption_set = "Create By Jielde"
        
    Case 3
        mcaption_set = "Scénario by Jielde"
        
    Case 4
        mcaption_set = "Programmation by"
        
    Case 5
        mcaption_set = "Jielde"
    
    Case 6
        mcaption_set = "Vous"
        
    Case 7
        mcaption_set = "Pour vbfrance.com"
        
    Case 8
        mcaption_set = "Echap pour sortir"

End Select
End Function
