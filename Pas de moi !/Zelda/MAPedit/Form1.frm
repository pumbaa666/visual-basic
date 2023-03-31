VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public Keyb As New keyboard
Private Sub Form_DblClick()
Ok = -1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Keyb.SetKeysDown KeyCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Keyb.SetKeysUp KeyCode
If KeyCode = vbKeyF1 Then Call map_SAVE
'If KeyCode = vbKeyF2 Then Ok = -1: Unloade: frm_param_map.Show

End Sub
Private Sub Form_Load()
Form1.Top = 0
Form1.Left = 0
Form1.Width = Screen.Width
Form1.Height = Screen.Height
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'system perso pour l'affichage d'un curseur
persoX = Int(x / Screen.TwipsPerPixelX)
persoY = Int(y / Screen.TwipsPerPixelY)
End Sub
