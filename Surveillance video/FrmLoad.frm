VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmLoad 
   Caption         =   "Calcul en cours"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   360
      Top             =   240
   End
   Begin ComctlLib.ProgressBar Bar 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Static vCntTimer As Boolean
    If vCntTimer = False Then
        Bar.Value = Bar.Value + 1
        If Bar.Value > 99 Then
            vCntTimer = True
        End If
    Else
        Bar.Value = Bar.Value - 1
        If Bar.Value > 99 Then
            vCntTimer = False
        End If
    End If
End Sub
