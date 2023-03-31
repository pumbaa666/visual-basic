VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Elf Beta 5 By Frecky "
   ClientHeight    =   4425
   ClientLeft      =   345
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3510
   Begin VB.CommandButton bs 
      Caption         =   "<--"
      Height          =   255
      Left            =   2640
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cl 
      Caption         =   "CLEAR"
      Height          =   255
      Left            =   2640
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Send 
      Caption         =   "SEND"
      Height          =   495
      Left            =   1800
      TabIndex        =   38
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox entry 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   260
      Left            =   480
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   480
      Width           =   1185
   End
   Begin VB.CommandButton space 
      Caption         =   " "
      Height          =   375
      Left            =   3000
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton elf 
      Caption         =   "ELF PWN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1250
      Left            =   480
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   4920
   End
   Begin VB.CommandButton rofl 
      Caption         =   "ROFL"
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton fy 
      BackColor       =   &H00404040&
      Caption         =   "FUCK YOU"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00404040&
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton noob 
      Caption         =   "NOOB"
      Height          =   375
      Left            =   1320
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton lmao 
      Caption         =   "LMAO"
      Height          =   375
      Left            =   360
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton wtf 
      Caption         =   "WTF"
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton lol 
      Caption         =   "LOL"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton stfu 
      Caption         =   "STFU"
      Height          =   375
      Left            =   360
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton z 
      Caption         =   "Z"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton y 
      Caption         =   "Y"
      Height          =   375
      Left            =   2280
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton x 
      Caption         =   "X"
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton w 
      Caption         =   "W"
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton v 
      Caption         =   "V"
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton u 
      Caption         =   "U"
      Height          =   375
      Left            =   840
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton t 
      Caption         =   "T"
      Height          =   375
      Left            =   480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton s 
      Caption         =   "S"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton r 
      Caption         =   "R"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton q 
      Caption         =   "Q"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton p 
      Caption         =   "P"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton o 
      Caption         =   "O"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton n 
      Caption         =   "N"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton m 
      Caption         =   "M"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton l 
      Caption         =   "L"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton k 
      Caption         =   "K"
      Height          =   375
      Left            =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cj 
      Caption         =   "J"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton ci 
      Caption         =   "I"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton ch 
      Caption         =   "H"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cg 
      Caption         =   "G"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cf 
      Caption         =   "F"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton ce 
      Caption         =   "E"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cd 
      Caption         =   "D"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cc 
      Caption         =   "C"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cb 
      Caption         =   "B"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton ca 
      Caption         =   "A"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.Label info 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   41
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label cmdX 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   36
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".::Emoticon Letter Flooder::."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4320
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstX, FirstY As Integer
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Private Function pl(lt As String)
If a = "" Then
a = lt
Call barrerbas
ElseIf b = "" Then
b = lt
ElseIf c = "" Then
c = lt
ElseIf d = "" Then
d = lt
ElseIf e = "" Then
e = lt
ElseIf f = "" Then
f = lt
ElseIf g = "" Then
g = lt
ElseIf h = "" Then
h = lt
ElseIf i = "" Then
i = lt
ElseIf j = "" Then
j = lt
Barrer
End If
entry.Text = entry.Text & lt
End Function
Private Sub bs_Click()
entry.Text = ""
If j <> "" Then
j = ""
entry.Text = a + b + c + d + e + f + g + h + i
ElseIf i <> "" Then
i = ""
entry.Text = a + b + c + d + e + f + g + h
ElseIf h <> "" Then
h = ""
entry.Text = a + b + c + d + e + f + g
ElseIf g <> "" Then
g = ""
entry.Text = a + b + c + d + e + f
ElseIf f <> "" Then
f = ""
entry.Text = a + b + c + d + e
ElseIf e <> "" Then
e = ""
entry.Text = a + b + c + d
ElseIf d <> "" Then
d = ""
entry.Text = a + b + c
ElseIf c <> "" Then
c = ""
entry.Text = a + b
ElseIf b <> "" Then
b = ""
entry.Text = a
ElseIf a <> "" Then
a = ""
barrerbas
End If
End Sub
Private Sub ca_Click()
Call pl("a")
End Sub
Private Sub cb_Click()
Call pl("b")
End Sub
Private Sub cc_Click()
Call pl("c")
End Sub
Private Sub cd_Click()
Call pl("d")
End Sub
Private Sub ce_Click()
Call pl("e")
End Sub
Private Sub cf_Click()
Call pl("f")
End Sub
Private Sub cg_Click()
Call pl("g")
End Sub
Private Sub ch_Click()
Call pl("h")
End Sub
Private Sub ci_Click()
Call pl("i")
End Sub
Private Sub cj_Click()
Call pl("j")
End Sub
Private Sub cl_Click()
a = ""
b = ""
c = ""
d = ""
e = ""
f = ""
g = ""
h = ""
i = ""
j = ""
entry.Text = ""
barrerbas
Barrer
End Sub
Private Sub elf_Click()
Call pl("e")
Call pl("l")
Call pl("f")
Call pl(" ")
Call pl("p")
Call pl("w")
Call pl("n")
tim
End Sub
Private Sub entry_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
entry.BackColor = &HFF8080
entry.BackColor = &H800000
End Sub
Private Sub fy_Click()
Call pl("f")
Call pl("u")
Call pl("c")
Call pl("k")
Call pl(" ")
Call pl("y")
Call pl("u")
Call pl("o")
tim
End Sub
Private Sub info_Click()
MsgBox "Elf made by Frecky, 2786184@nims.csrs.qc.ca" + vbCrLf + "E-mail me for Bug or things to add" + vbCrLf + "Beta Tester: Xantim" + vbCrLf + vbCrLf + "To use, Select Pre-make text or" + vbCrLf + "select letter and click send then go" + vbCrLf + "in your msn window and let Elf Do the job!", vbInformation + vbOKOnly + vbMsgBoxSetForeground, ".::About Elf::."
End Sub
Private Sub k_Click()
Call pl("k")
End Sub
Private Sub l_Click()
Call pl("l")
End Sub
Private Sub lmao_Click()
Call pl("l")
Call pl("m")
Call pl("a")
Call pl("o")
tim
End Sub
Private Sub lol_Click()
Call pl("l")
Call pl("o")
Call pl("l")
tim
End Sub
Private Sub m_Click()
Call pl("m")
End Sub
Private Sub n_Click()
Call pl("n")
End Sub
Private Sub noob_Click()
Call pl("n")
Call pl("o")
Call pl("o")
Call pl("b")
tim
End Sub
Private Sub o_Click()
Call pl("o")
End Sub
Private Sub p_Click()
Call pl("p")
End Sub
Private Sub q_Click()
Call pl("q")
End Sub
Private Sub r_Click()
Call pl("r")
End Sub
Private Sub rofl_Click()
Call pl("r")
Call pl("o")
Call pl("f")
Call pl("l")
Timer1.Enabled = True
End Sub
Private Sub s_Click()
Call pl("s")
End Sub
Private Sub send_Click()
tim
End Sub
Private Sub space_Click()
Call pl(" ")
End Sub
Private Sub stfu_Click()
Call pl("s")
Call pl("t")
Call pl("f")
Call pl("u")
Timer1.Enabled = True
End Sub
Private Sub t_Click()
Call pl("t")
End Sub
Private Sub Timer1_Timer()
Timer2.Enabled = True
Timer1.Enabled = False
End Sub
Private Sub Timer2_Timer()
Call check
End Sub
Private Sub u_Click()
Call pl("u")
End Sub
Private Sub v_Click()
Call pl("v")
End Sub
Private Sub w_Click()
Call pl("w")
End Sub
Private Sub wtf_Click()
Call pl("w")
Call pl("t")
Call pl("f")
Timer1.Enabled = True
End Sub
Private Sub x_Click()
Call pl("x")
End Sub
Private Sub y_Click()
Call pl("y")
End Sub
Private Sub z_Click()
Call pl("z")
End Sub
Private Function Sa()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sb()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sc()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sd()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function se()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sf()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sg()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sh()
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function si()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sj()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sk()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sl()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sm()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:@:S:@:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sn()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:@:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function so()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sp()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sq()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:@:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:@:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sr()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function ss()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function st()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function su()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sv()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sw()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:@:S"
SendKeys "+~"
SendKeys ":S:S:@:S:@:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:@:@:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sx()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sy()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:S:S:S:@:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sz()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:@:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:@:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:@:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:@:@:@:@:@:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function sspace()
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "+~"
SendKeys ":S:S:S:S:S:S:S:S"
SendKeys "~"
End Function
Private Function Barrer()
ca.Enabled = Not ca.Enabled
cb.Enabled = Not cb.Enabled
cc.Enabled = Not cc.Enabled
cd.Enabled = Not cd.Enabled
ce.Enabled = Not ce.Enabled
cf.Enabled = Not cf.Enabled
cg.Enabled = Not cg.Enabled
ch.Enabled = Not ch.Enabled
ci.Enabled = Not ci.Enabled
cj.Enabled = Not cj.Enabled
k.Enabled = Not k.Enabled
l.Enabled = Not l.Enabled
m.Enabled = Not m.Enabled
n.Enabled = Not n.Enabled
o.Enabled = Not o.Enabled
p.Enabled = Not p.Enabled
q.Enabled = Not q.Enabled
r.Enabled = Not r.Enabled
s.Enabled = Not s.Enabled
t.Enabled = Not t.Enabled
u.Enabled = Not u.Enabled
v.Enabled = Not v.Enabled
w.Enabled = Not w.Enabled
x.Enabled = Not x.Enabled
y.Enabled = Not y.Enabled
z.Enabled = Not z.Enabled
space.Enabled = Not space.Enabled
End Function
Private Function barrerbas()
fy.Enabled = Not fy.Enabled
stfu.Enabled = Not stfu.Enabled
lol.Enabled = Not lol.Enabled
wtf.Enabled = Not wtf.Enabled
lmao.Enabled = Not lmao.Enabled
noob.Enabled = Not noob.Enabled
rofl.Enabled = Not rofl.Enabled
elf.Enabled = Not elf.Enabled
End Function
Private Function check()
If a <> "" Then
checker a
a = ""
ElseIf b <> "" Then
checker b
b = ""
ElseIf c <> "" Then
checker c
c = ""
ElseIf d <> "" Then
checker d
d = ""
ElseIf e <> "" Then
checker e
e = ""
ElseIf f <> "" Then
checker f
f = ""
ElseIf g <> "" Then
checker g
g = ""
ElseIf h <> "" Then
checker h
h = ""
ElseIf i <> "" Then
checker i
i = ""
ElseIf j <> "" Then
checker j
j = ""
Else
Timer2.Enabled = False
Send.Enabled = True
Barrer
barrerbas
cl.Enabled = True
bs.Enabled = True
End If
End Function
Private Function tim()
Barrer
fy.Enabled = False
stfu.Enabled = False
lol.Enabled = False
wtf.Enabled = False
lmao.Enabled = False
noob.Enabled = False
rofl.Enabled = False
elf.Enabled = False
Send.Enabled = False
entry.Text = ""
Timer1.Enabled = True
cl.Enabled = False
bs.Enabled = False
End Function
Private Sub cmdX_Click()
Unload Me
End Sub
Private Sub cmdX_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdX.ForeColor = &HC00000
End Sub
Private Sub info_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
info.ForeColor = &HC00000
End Sub
Private Sub title_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    FirstX = x
    FirstY = y
End If
End Sub
Private Sub title_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Me.Left = Me.Left + (x - FirstX)
    Me.Top = Me.Top + (y - FirstY)
End If
cmdX.BackColor = &HFF8080
info.BackColor = &HFF8080
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
entry.BackColor = &HFF8080
cmdX.ForeColor = &HFFFFFF
End Sub
Private Function checker(lett As String)
Select Case lett
Case "a"
Sa
Case "b"
sb
Case "c"
sc
Case "d"
sd
Case "e"
se
Case "f"
sf
Case "g"
sg
Case "h"
sh
Case "i"
si
Case "j"
sj
Case "k"
sk
Case "l"
sl
Case "m"
sm
Case "n"
sn
Case "o"
so
Case "p"
sp
Case "q"
sq
Case "r"
sr
Case "s"
ss
Case "t"
st
Case "u"
su
Case "v"
sv
Case "w"
sw
Case "x"
sx
Case "y"
sy
Case "z"
sz
Case " "
sspace
End Select
End Function
