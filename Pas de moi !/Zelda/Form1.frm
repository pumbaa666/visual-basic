VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer bomb_reload_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   720
   End
   Begin VB.Timer Sword_timer 
      Enabled         =   0   'False
      Interval        =   510
      Left            =   6000
      Top             =   0
   End
   Begin VB.Timer Anim_timer 
      Interval        =   250
      Left            =   6000
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Keyb As New keyboard

Private Sub Anim_timer_Timer()
   If Mod_DD.animtile_index < 2 Then Mod_DD.animtile_index = Mod_DD.animtile_index + 1 Else Mod_DD.animtile_index = 0
End Sub

Private Sub bomb_reload_timer_Timer()
   bomb_reload_timer.Enabled = False
End Sub

Private Sub Form_DblClick()
   Ok = -1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Keyb.SetKeysDown KeyCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Keyb.SetKeysUp KeyCode
End Sub
Private Sub Form_Load()
    Form1.Top = 0
    Form1.Left = 0
    Form1.Width = Screen.Width
    Form1.Height = Screen.Height
End Sub
Private Sub Sword_timer_Timer()
    Sword_timer.Enabled = False
    sword_state = 0
End Sub
