Attribute VB_Name = "Mod_INVENTAIRE"
'---------------------------------------------------------------------------------------
' Module    : Mod_INVENTAIRE
' DateTime  : 30/12/2004 09:45
' Author    : Gwenael
' Bon la c'est que des déclarations de variables, ou de types
'---------------------------------------------------------------------------------------

Public Type bomb
nb As Integer
nbposees As Integer
End Type

Public Type bomb_posee
x As Single
y As Single
timer As Integer
End Type

Public Type H_Armes
   epee As Byte
   bombes As bomb
   arc As Boolean
End Type

Public Type H_Objets
   rubis As Integer
   bottes As Boolean
   palmes As Boolean
End Type

Public Type H_statut
vie As Single
vitesse As Single
End Type

Public Type HERO
Armes As H_Armes
Objets As H_Objets
statut As H_statut
End Type

'Tiré de :
' RPG ENGINE FOR WINDOWS - DECLARATIONS
' (C) 2003, Fling-master

Public Type BulletType
 x       As Integer            'X and Y position of bullet
 y       As Integer
 Dir     As Integer            'Direction it's headed in
 speed   As Integer            'ça c'est de moi
 surf    As DirectDrawSurface7 'ça aussi
 Damage  As Integer            'Damage dealt to whatever it hits
End Type
'------------------------------------------------------------------
Public Heros As HERO

Public Type JEU
statut As String
End Type

Public JEU As JEU

Public bomb_posee(5) As bomb_posee
Public bomb_counter As Integer
Public bomb_timer(5) As Integer
Public bomb_explosion_timer(5) As Integer
