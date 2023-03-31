VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Veuillez patienter"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Timer ClkProgress 
      Interval        =   50
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "En attente d'un joueur"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "FrmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClkProgress_Timer()
Static vCount As Integer
    ProgressBar.Value = vCount
    vCount = vCount + 1
    If vCount = 100 Then
        vCount = 0
    End If
End Sub

Private Sub CmdAnnuler_Click()
    ClkProgress.Enabled = False
    FrmOptMulti.Wsk.Close
    FrmWait.Hide
End Sub
