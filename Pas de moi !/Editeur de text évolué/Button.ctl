VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label LabelClick 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label LabelCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   585
   End
   Begin VB.Image ImageButton 
      Height          =   255
      Left            =   840
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   600
      X2              =   2160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   2160
      X2              =   2160
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   600
      X2              =   2160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   600
      X2              =   600
      Y1              =   600
      Y2              =   1080
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Variables générales
Dim Down As Boolean

'Déclaration des variables de stockage des propriétés
Enum ButtonAppearance
    StyleFlat = 0
    Style3D = 1
End Enum
Dim m_Appearance As ButtonAppearance
Dim m_HighLightColor3D As OLE_COLOR
Dim m_LowLightColor3D As OLE_COLOR
Dim m_ColorFlat As OLE_COLOR
Dim m_ForeColor As OLE_COLOR

'Déclaration des constantes pour les propriétés de base
Const b_Appearance = 1
Const b_HighLightColor3D = &HE0E0E0
Const b_LowLightColor3D = &H80000008
Const b_ColorFlat = &H80000008
Const b_BackColor = &H8000000F
Const b_Caption = "Button"
Const b_ForeColor = &H80000012
Const b_Enabled = True

'Déclaration des évènements
Event Click()
Private Sub LabelClick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Down = False Then
    If m_Appearance = Style3D Then
        Line1.BorderColor = m_LowLightColor3D
        Line2.BorderColor = m_LowLightColor3D
        Line3.BorderColor = m_HighLightColor3D
        Line4.BorderColor = m_HighLightColor3D
    Else
        LabelCaption.Move LabelCaption.Left + 10, LabelCaption.Top + 10
    End If
    Down = True
End If

End Sub

Private Sub LabelClick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Down = True Then
    If m_Appearance = Style3D Then
        Line1.BorderColor = m_HighLightColor3D
        Line2.BorderColor = m_HighLightColor3D
        Line3.BorderColor = m_LowLightColor3D
        Line4.BorderColor = m_LowLightColor3D
    Else
        LabelCaption.Move LabelCaption.Left - 10, LabelCaption.Top - 10
    End If
    RaiseEvent Click
    Down = False
End If

End Sub

Private Sub UserControl_Resize()

'On vérifie si on ce n'est pas trop petit
If Height < 100 Then Height = 100
If Width < 100 Then Width = 100

'On redimensionne bien les contrôles
Line1.X1 = 0
Line1.X2 = 0
Line1.Y1 = 0
Line1.Y2 = Height - 10
Line2.X1 = 0
Line2.X2 = Width - 10
Line2.Y1 = 0
Line2.Y2 = 0
Line3.X1 = Width - 10
Line3.X2 = Width - 10
Line3.Y1 = 0
Line3.Y2 = Height - 10
Line4.X1 = 0
Line4.X2 = Width - 10
Line4.Y1 = Height - 10
Line4.Y2 = Height - 10
LabelCaption.Move (Width - LabelCaption.Width) / 2, (Height - LabelCaption.Height) / 2
ImageButton.Move 0, 0, Width - 10, Height - 10
LabelClick.Move 0, 0, Width, Height

End Sub

Public Property Get HighLightColor3D() As OLE_COLOR

HighLightColor3D = m_HighLightColor3D

If m_Appearance = Style3D Then
Line1.BorderColor = m_HighLightColor3D
Line2.BorderColor = m_HighLightColor3D
End If

End Property

Public Property Let HighLightColor3D(ByVal New_HighLightColor3D As OLE_COLOR)

m_HighLightColor3D = New_HighLightColor3D
PropertyChanged "HighLightColor3D"
  
End Property

Public Property Get LowLightColor3D() As OLE_COLOR

LowLightColor3D = m_LowLightColor3D

If m_Appearance = Style3D Then
Line3.BorderColor = m_LowLightColor3D
Line4.BorderColor = m_LowLightColor3D
End If

End Property

Public Property Let LowLightColor3D(ByVal New_LowLightColor3D As OLE_COLOR)

m_LowLightColor3D = New_LowLightColor3D
PropertyChanged "LowLightColor3D"

End Property

Public Property Get ColorFlat() As OLE_COLOR

ColorFlat = m_ColorFlat

If m_Appearance = StyleFlat Then
Line1.BorderColor = m_ColorFlat
Line2.BorderColor = m_ColorFlat
Line3.BorderColor = m_ColorFlat
Line4.BorderColor = m_ColorFlat
End If

End Property

Public Property Let ColorFlat(ByVal New_ColorFlat As OLE_COLOR)

m_ColorFlat = New_ColorFlat
PropertyChanged "ColorFlat"

End Property

Public Property Get Appearance() As ButtonAppearance

Appearance = m_Appearance

If m_Appearance = Style3D Then
    Line1.BorderColor = m_HighLightColor3D
    Line2.BorderColor = m_HighLightColor3D
    Line3.BorderColor = m_LowLightColor3D
    Line4.BorderColor = m_LowLightColor3D
Else
    Line1.BorderColor = m_ColorFlat
    Line2.BorderColor = m_ColorFlat
    Line3.BorderColor = m_ColorFlat
    Line4.BorderColor = m_ColorFlat
End If

End Property

Public Property Let Appearance(ByVal NewAppearance As ButtonAppearance)

m_Appearance = NewAppearance
PropertyChanged "Appearance"

End Property

Public Property Get BackColor() As OLE_COLOR

BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

UserControl.BackColor = New_BackColor
PropertyChanged "BackColor"

End Property

Public Property Get Caption() As String

Caption = LabelCaption.Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

LabelCaption.Caption = New_Caption
PropertyChanged "Caption"

End Property

Public Property Get Pic() As IPictureDisp

Set Pic = ImageButton.Picture

End Property

Public Property Set Pic(ByVal New_Pic As IPictureDisp)

Set ImageButton.Picture = New_Pic
PropertyChanged "Pic"

End Property

Public Property Get ForeColor() As OLE_COLOR

ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal Color As OLE_COLOR)

m_ForeColor = Color
PropertyChanged "ForeColor"

If UserControl.Enabled = True Then LabelCaption.ForeColor = m_ForeColor

End Property

Public Property Get Enabled() As Boolean

Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

UserControl.Enabled = New_Enabled
PropertyChanged "Enabled"

If UserControl.Enabled = False Then
    LabelCaption.ForeColor = &H818181
Else
    LabelCaption.ForeColor = m_ForeColor
End If

End Property

Private Sub UserControl_InitProperties()

'Initialisation des propriétés
m_Appearance = b_Appearance
m_HighLightColor3D = b_HighLightColor3D
m_LowLightColor3D = b_LowLightColor3D
m_ColorFlat = b_ColorFlat
UserControl.BackColor = b_BackColor
LabelCaption.Caption = b_Caption
Set ImageButton.Picture = Nothing
m_ForeColor = b_ForeColor
UserControl.Enabled = b_Enabled

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'Lecture des propriétés
m_Appearance = PropBag.ReadProperty("Appearance", b_Appearance)
m_HighLightColor3D = PropBag.ReadProperty("HighLightColor3D", b_HighLightColor3D)
m_LowLightColor3D = PropBag.ReadProperty("LowLightColor3D", b_LowLightColor3D)
m_ColorFlat = PropBag.ReadProperty("ColorFlat", b_ColorFlat)
UserControl.BackColor = PropBag.ReadProperty("BackColor", b_BackColor)
LabelCaption.Caption = PropBag.ReadProperty("Caption", b_Caption)
Set ImageButton.Picture = PropBag.ReadProperty("Pic", Nothing)
m_ForeColor = PropBag.ReadProperty("ForeColor", b_ForeColor)
UserControl.Enabled = PropBag.ReadProperty("Enabled", b_Enabled)

'Chargement des propriétés
If UserControl.Enabled = False Then
    LabelCaption.ForeColor = &HE0E0E0
Else
    LabelCaption.ForeColor = m_ForeColor
End If
If m_Appearance = Style3D Then
    Line1.BorderColor = m_HighLightColor3D
    Line2.BorderColor = m_HighLightColor3D
    Line3.BorderColor = m_LowLightColor3D
    Line4.BorderColor = m_LowLightColor3D
Else
    Line1.BorderColor = m_ColorFlat
    Line2.BorderColor = m_ColorFlat
    Line3.BorderColor = m_ColorFlat
    Line4.BorderColor = m_ColorFlat
End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'Enregistrement des propriétés
Call PropBag.WriteProperty("Appearance", m_Appearance, b_Appearance)
Call PropBag.WriteProperty("HighLightColor3D", m_HighLightColor3D, b_HighLightColor3D)
Call PropBag.WriteProperty("LowLightColor3D", m_LowLightColor3D, b_LowLightColor3D)
Call PropBag.WriteProperty("ColorFlat", m_ColorFlat, b_ColorFlat)
Call PropBag.WriteProperty("BackColor", UserControl.BackColor, b_BackColor)
Call PropBag.WriteProperty("Caption", LabelCaption.Caption, b_Caption)
Call PropBag.WriteProperty("Pic", ImageButton.Picture, Nothing)
Call PropBag.WriteProperty("ForeColor", LabelCaption.ForeColor, b_ForeColor)

End Sub
