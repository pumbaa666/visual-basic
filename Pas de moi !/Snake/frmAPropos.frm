VERSION 5.00
Begin VB.Form frmAPropos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   0
      Picture         =   "frmAPropos.frx":0000
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label lblDateInfo 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblVersionInfo 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblAuteurInfo 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblDate 
      Caption         =   "Date :"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version :"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblAuteur 
      Caption         =   "Auteur :"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAPropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    frmAPropos.Hide
End Sub

Private Sub Form_Load()
    lblAuteurInfo.Caption = "Blanc Siméon"
    lblVersionInfo.Caption = "Version 1.0"
    lblDateInfo.Caption = "5 mai 2004"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
