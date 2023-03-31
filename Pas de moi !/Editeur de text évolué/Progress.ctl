VERSION 5.00
Begin VB.UserControl Progress 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   ScaleHeight     =   3135
   ScaleWidth      =   4545
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   2640
      X2              =   2640
      Y1              =   230
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   720
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   2630
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   720
      Y1              =   240
      Y2              =   470
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Déclaration des variables de stockage des propriétés
Enum ProgressAppearance
    StyleFlat = 0
    Style3D = 1
End Enum
Dim m_Appearance As ProgressAppearance
Dim m_HighLightColor3D As OLE_COLOR
Dim m_LowLightColor3D As OLE_COLOR
Dim m_ColorFlat As OLE_COLOR
Dim m_Value As Single
Dim m_Maxi As Single
Dim m_Mini As Single
Dim m_ColorBar As OLE_COLOR

'Déclaration des constantes pour les propriétés de base
Const b_Appearance = 1
Const b_HighLightColor3D = &H80000008
Const b_LowLightColor3D = &HE0E0E0
Const b_ColorFlat = &H80000008
Const b_Mini = 0
Const b_Maxi = 100
Const b_Value = 50
Const b_ColorBar = &H80FF&
Const b_BackColor = &H8000000F

Private Sub UserControl_Resize()

'On vérifie si on ce n'est pas trop petit
If Height < 255 Then Height = 255
If Width < 500 Then Width = 500

'On redimensionne bien les dessins
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

Public Property Get Appearance() As ProgressAppearance

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

Public Property Let Appearance(ByVal NewAppearance As ProgressAppearance)

m_Appearance = NewAppearance
PropertyChanged "Appearance"

End Property

Public Property Get Value() As Single

Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As Single)

If NewValue > Maxi Or NewValue < Mini Then MsgBox "Value doit se trouver entre Mini et Maxi.", vbExclamation, "Erreur": Exit Property

m_Value = NewValue
PropertyChanged "Value"

Line (10, Height - 30)-(Width - 20, 10), BackColor, BF
If Value <> 0 Then Line (10, Height - 30)-((Value * ((Width - 20) / (Maxi - Mini))) - 20, 10), ColorBar, BF

End Property

Public Property Get Maxi() As Single

Maxi = m_Maxi

End Property

Public Property Let Maxi(ByVal NewMaxi As Single)

If NewMaxi < m_Mini Then MsgBox "Maxi ne peut être inférieur à Mini.", vbExclamation, "Erreur": Exit Property
If NewMaxi < m_Value Then MsgBox "Maxi ne peut être inférieur à Value.", vbExclamation, "Erreur": Exit Property

m_Maxi = NewMaxi
PropertyChanged "Maxi"

End Property

Public Property Get Mini() As Single

Mini = m_Mini

End Property

Public Property Let Mini(ByVal NewMini As Single)

If NewMini < 0 Then MsgBox "Mini ne peut être inférieur à 0.", vbExclamation, "Erreur": Exit Property
If NewMini > m_Maxi Then MsgBox "Mini ne peut être supérieur à Maxi.", vbExclamation, "Erreur": Exit Property
If NewMini > m_Value Then MsgBox "Mini ne peut être supérieur à Value.", vbExclamation, "Erreur": Exit Property

m_Mini = NewMini
PropertyChanged "Mini"

End Property

Public Property Get ColorBar() As OLE_COLOR

ColorBar = m_ColorBar

End Property

Public Property Let ColorBar(ByVal New_ColorBar As OLE_COLOR)

m_ColorBar = New_ColorBar
PropertyChanged "ColorBar"

End Property

Public Property Get BackColor() As OLE_COLOR

BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

UserControl.BackColor = New_BackColor
PropertyChanged "BackColor"

End Property

Private Sub UserControl_InitProperties()

'Initialisation des propriétés
m_Appearance = b_Appearance
m_HighLightColor3D = b_HighLightColor3D
m_LowLightColor3D = b_LowLightColor3D
m_ColorFlat = b_ColorFlat
m_Value = b_Value
m_Maxi = b_Maxi
m_Mini = b_Mini
UserControl.BackColor = b_BackColor

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'Lecture des propriétés
m_Appearance = PropBag.ReadProperty("Appearance", b_Appearance)
m_HighLightColor3D = PropBag.ReadProperty("HighLightColor3D", b_HighLightColor3D)
m_LowLightColor3D = PropBag.ReadProperty("LowLightColor3D", b_LowLightColor3D)
m_ColorFlat = PropBag.ReadProperty("ColorFlat", b_ColorFlat)
m_Value = PropBag.ReadProperty("Value", b_Value)
m_Maxi = PropBag.ReadProperty("Maxi", b_Maxi)
m_Mini = PropBag.ReadProperty("Mini", b_Mini)
m_ColorBar = PropBag.ReadProperty("ColorBar", b_ColorBar)
UserControl.BackColor = PropBag.ReadProperty("BackColor", b_BackColor)

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
If Value <> 0 Then Line (10, Height - 30)-((Value * ((Width - 20) / (Maxi - Mini))) - 20, 10), m_ColorBar, BF

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'Enregistrement des propriétés
Call PropBag.WriteProperty("Appearance", m_Appearance, b_Appearance)
Call PropBag.WriteProperty("HighLightColor3D", m_HighLightColor3D, b_HighLightColor3D)
Call PropBag.WriteProperty("LowLightColor3D", m_LowLightColor3D, b_LowLightColor3D)
Call PropBag.WriteProperty("ColorFlat", m_ColorFlat, b_ColorFlat)
Call PropBag.WriteProperty("Value", m_Value, b_Value)
Call PropBag.WriteProperty("Maxi", m_Maxi, b_Maxi)
Call PropBag.WriteProperty("Mini", m_Mini, b_Mini)
Call PropBag.WriteProperty("ColorBar", m_ColorBar, b_ColorBar)
Call PropBag.WriteProperty("BackColor", UserControl.BackColor, b_BackColor)

End Sub
