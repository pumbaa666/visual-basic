VERSION 5.00
Begin VB.UserControl CoolBar 
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   PropertyPages   =   "CoolBar.ctx":0000
   ScaleHeight     =   4170
   ScaleWidth      =   7365
   Begin VB.Image ImageButton 
      Height          =   480
      Index           =   1
      Left            =   30
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   2760
      Y1              =   550
      Y2              =   550
   End
   Begin VB.Line Line3 
      X1              =   2760
      X2              =   2760
      Y1              =   0
      Y2              =   550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   2760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   480
   End
End
Attribute VB_Name = "CoolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Variables générales
Dim Emplacement As Integer
Dim Num As Integer
Dim TabKey() As String
Dim Down As Boolean

'Déclaration des variables de stockage des propriétés
Enum CoolBarAppearance
    StyleFlat = 0
    Style3D = 1
End Enum
Dim m_Appearance As CoolBarAppearance
Dim m_HighLightColor3D As OLE_COLOR
Dim m_LowLightColor3D As OLE_COLOR
Dim m_ColorFlat As OLE_COLOR

'Déclaration des constantes pour les propriétés de base
Const b_Appearance = 1
Const b_HighLightColor3D = &HE0E0E0
Const b_LowLightColor3D = &H80000008
Const b_ColorFlat = &H80000008
Const b_BackColor = &H8000000F
Const b_Enabled = True

'Déclaration des évènements
Event ButtonClick(Key As String)

Public Sub AddButton(Pic As Integer, ImgList As ImageList, Key As String, Tips As String)

'Sub pour ajouter un bouton en temps réel
'(le seul moyen d'en ajouter pour le moment)

Num = Num + 1
ReDim Preserve TabKey(Num)
TabKey(Num) = Key

If Num <> 1 Then
    Load ImageButton(Num)
    Emplacement = Emplacement + 550
    ImageButton(Num).Move Emplacement
End If

ImageButton(Num).Picture = ImgList.ListImages(Pic).Picture
ImageButton(Num).ToolTipText = Tips
ImageButton(Num).Visible = True

End Sub

Private Sub ImageButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Down = False Then
    ImageButton(Index).Move ImageButton(Index).Left + 20, ImageButton(Index).Top + 20
    Down = True
End If

End Sub

Private Sub ImageButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Down = True Then
    ImageButton(Index).Move ImageButton(Index).Left - 20, ImageButton(Index).Top - 20
    RaiseEvent ButtonClick(TabKey(Index))
    Down = False
End If

End Sub

Private Sub UserControl_Resize()

'Le contrôle est vérouillé en dimensions
Height = 550
If Width < 550 Then Width = 550

'On replace bien les contrôles
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

Public Property Get Appearance() As CoolBarAppearance

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

Public Property Let Appearance(ByVal NewAppearance As CoolBarAppearance)

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

Public Property Get Enabled() As Boolean

Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

UserControl.Enabled = New_Enabled
PropertyChanged "Enabled"

End Property

Private Sub UserControl_InitProperties()

'Initialisation des propriétés
m_Appearance = b_Appearance
m_HighLightColor3D = b_HighLightColor3D
m_LowLightColor3D = b_LowLightColor3D
m_ColorFlat = b_ColorFlat
UserControl.BackColor = b_BackColor
UserControl.Enabled = b_Enabled

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'Lecture des propriétés
m_Appearance = PropBag.ReadProperty("Appearance", b_Appearance)
m_HighLightColor3D = PropBag.ReadProperty("HighLightColor3D", b_HighLightColor3D)
m_LowLightColor3D = PropBag.ReadProperty("LowLightColor3D", b_LowLightColor3D)
m_ColorFlat = PropBag.ReadProperty("ColorFlat", b_ColorFlat)
UserControl.BackColor = PropBag.ReadProperty("BackColor", b_BackColor)
UserControl.Enabled = PropBag.ReadProperty("Enabled", b_Enabled)

'Chargement des propriétés
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
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, b_Enabled)

End Sub
