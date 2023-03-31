VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Chenillard"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbStyle 
      Height          =   315
      ItemData        =   "FrmMain.frx":0000
      Left            =   4200
      List            =   "FrmMain.frx":0010
      TabIndex        =   3
      Text            =   "Style du chenillard"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Sto&p"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer ClkLed 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4320
      Top             =   1080
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   840
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape ShpLed 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCount As Integer

Private Sub ClkLed_Timer()
Static vTest As Boolean
    If CmbStyle.Text = 1 Then
        If vCount = 8 Then
            vCount = 0
            AllLed ("&H80000004")
        End If
        ShpLed(vCount).FillColor = &HFF&
        vCount = vCount + 1
    ElseIf CmbStyle = 2 Then
        If vCount = 8 Then
            vTest = True
            vCount = 7
        ElseIf vCount = -1 Then
            vTest = False
            vCount = 0
        End If
        
        If vTest = False Then
            ShpLed(vCount).FillColor = &HFF&
            vCount = vCount + 1
        Else
            ShpLed(vCount).FillColor = &H80000004
            vCount = vCount - 1
        End If
    End If
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub CmdStart_Click()
    AllLed ("&H80000004")
    If IsNumeric(CmbStyle.Text) = False Then
        MsgBox "Le n° du chenillard n'est pas valide!", vbCritical, "Erreur"
    Else
        If CmbStyle.Text = 1 Then
            vCount = 0
        ElseIf CmbStyle = 2 Then
            vCount = 0
        End If
        ClkLed.Enabled = True
    End If
End Sub

Private Sub CmdStop_Click()
    ClkLed.Enabled = False
    AllLed ("&H80000004")
End Sub

Function AllLed(ByVal vCouleur As String)
Dim vCntAll As Integer
    For vCntAll = 0 To 7
        ShpLed(vCntAll).FillColor = vCouleur
    Next
End Function
