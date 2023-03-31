Attribute VB_Name = "Mod_DD"
'---------------------------------------------------------------------------------------
' Module    : Mod_DD
' DateTime  : 16/06/2004 11:44
' Author    : Gwenael
 
'Merci a AbeLeMudokon pour sa fonction AfficherImage très utile(ça m'a évité de me
'prendre la tete sur cette partie de code plutot chiante ;-))
 
'Je tiens a préciser que toutes les fonction de GAMMA ne sont pas de moi, mais du site
'http://www.dx4vb.da.ru . Merci a eux pour leur tutorial!

'DESOLE POUR LE MANQUE DE COMMENTAIRES (LA FLEMME)
 
'SVP LAISSEZ DES COMMENTAIRES SUR VBFRANCE CE SERAIT SYMPA(ET CA PERMET DE FAIRE
'PROGRESSER LA SCOURCE!)

'Si vous n'avez rien a faire quand vous allez sur Internet, allez sur :
'http://www.sitealpha.fr.fm/
'---------------------------------------------------------------------------------------
Public DX As New DirectX7
Public DD As DirectDraw7

'directSound
Public DS As DirectSound
Public DSWaveFormat As WAVEFORMATEX    'format Wave
'sons
Public BuffSons() As DirectSoundBuffer   'buffers directSound
Public DescSons() As DSBUFFERDESC        'descripteurs de buffer directSound
'DirectMusic
Dim perf As DirectMusicPerformance
' Des informations sur  le fichier MIDI peuvent être recueillis dans cette variable
Dim seg As DirectMusicSegment
' Stockage du fichier MIDI
Dim segstate As DirectMusicSegmentState
' Statut du fichier charger
Dim loader As DirectMusicLoader
' Chargement d'un fichier MIDI dans une variable DirectMusicSegment
Public DM_FileName As String
Public Primary As DirectDrawSurface7    'Surface primaire visible a l'écran
Public Backbuffer As DirectDrawSurface7 'Surface brouillon invisible
Public Ok As Integer
Public ColorKey As DDCOLORKEY

'=============================================================================================

Public Tile(200) As DirectDrawSurface7
Public Tileddsd As DDSURFACEDESC2
Public Perso(1 To 4) As DirectDrawSurface7

Public perso_sword_H(8) As DirectDrawSurface7
Public perso_sword_B(8) As DirectDrawSurface7
Public perso_sword_L(8) As DirectDrawSurface7
Public perso_sword_R(8) As DirectDrawSurface7
Public Perso_swordddsd_H(8) As DDSURFACEDESC2
Public Perso_swordddsd_B(8) As DDSURFACEDESC2
Public Perso_swordddsd_L(8) As DDSURFACEDESC2
Public Perso_swordddsd_R(8) As DDSURFACEDESC2

Public Titre As DirectDrawSurface7
Public Titreddsd As DDSURFACEDESC2

Public Persoddsd(1 To 4) As DDSURFACEDESC2

Public charsurf(4) As DirectDrawSurface7
Public charsurfddsd(4) As DDSURFACEDESC2
Public OBJsurf(1) As DirectDrawSurface7
Public OBJsurfddsd(1) As DDSURFACEDESC2
Public ENNEMIsurf(3) As DirectDrawSurface7
Public ENNEMIsurfddsd(3) As DDSURFACEDESC2

Public bombsurf As DirectDrawSurface7
Public bombsurf2 As DirectDrawSurface7

Public explosion_surf As DirectDrawSurface7
Public explosion_surfddsd As DDSURFACEDESC2

Public bombsurfddsd As DDSURFACEDESC2

Public PosMondeX
Public PosMondeY

Public tileON_bottomright
Public tileON_topleft
Public tileON_bottomleft
Public tileON_topright

Public StopJeu As Boolean

'---- GAMMA correction
Dim GammaControler As DirectDrawGammaControl
Dim GammaRamp As DDGAMMARAMP
Dim OriginalRamp As DDGAMMARAMP
Dim GammaSupport As Boolean
Dim CurrRed As Double, CurrGreen As Double, CurrBlue As Double

Public screenwidth
Public screenheight

Public Map()
Public Type param_map
 nom As String
End Type
Public Type WarpZone
 index As Integer
 x As String
 y As String
 destX As String
 destY As String
 DestMap As String
End Type
Public Type char
 index As Integer
 x As String
 y As String
 Txt As String
 img As String
 imgddsd As String
End Type
Public Type OBJ
 x As String
 y As String
 img As String
 type As String
End Type
Public Type ENNEMI
 x As String
 y As String
 type As String
End Type

Public param_map As param_map
Public longueurMapX As Integer
Public longueurMapY As Integer

Public camlock_H As Boolean
Public camlock_B As Boolean
Public camlock_L As Boolean
Public camlock_R As Boolean


Public WarpZone() As WarpZone
Public nbrWarp As Integer
Public nbrChar As Integer
Public nbrOBJ As Integer
Public nbrENNEMI As Integer
Public char() As char
Public OBJ() As OBJ
Public ENNEMI(50) As ENNEMI

Public animtile_index As Integer
Public perso_index As Integer
Public persoX As Integer
Public persoY As Integer
Public perso_move As Boolean
Public perso_dir As String

Public msg As String
Public msgcount As Integer
Public sword_state As Integer
Public sword_possible As Boolean
Public bomb_possible As Boolean
Public anim_sword As Single
'---------------------------------------------------------------------------------------
' Procedure : Main
' DateTime  : 05/12/2004 20:26
' Author    : Gwenael
'---------------------------------------------------------------------------------------
Sub Main()

screenwidth = 800
screenheight = 600
  persoX = 320
  persoY = 224
perso_index = 1
perso_move = True

LoadDS
LoadJEU
Form2.Hide
Unload Form2

Do

Backbuffer.BltColorFill ddRect(0, 0, 0, 0), 0
If JEU.statut = "" Then
afficheMAP
End If
If Heros.statut.vie < 0 Then Heros.statut.vie = 0
If Heros.statut.vie = 0 Then JEU.statut = "gameover"
If JEU.statut = "menuprincipal" Then
AfficherImage Titre, Titreddsd, 100, 100, ddRect(0, 0, 0, 0)
Backbuffer.DrawText 50, 20, "Projet VB_Zelda", False
If Form1.Keyb.SpaceKey Then JEU.statut = ""
End If

If JEU.statut = "gameover" Then
BuffSons(2).Play (DSBPLAY_DEFAULT)
Heros.Armes.epee = 0
Backbuffer.DrawText 250, 50, "G A M E      O V E R", False
End If
If JEU.statut = "" Then
'------------------------------------------------
Check_Collision
Check_Warp
move_collision
move
'------------------------------------------------

If msg <> "" Then
Backbuffer.DrawText 50, 1, msg, False
perso_move = False
If msgcount < 20 Then msgcount = msgcount + 1
Else: perso_move = True
End If

If Form1.Keyb.SpaceKey = True And msgcount >= 20 Then msg = ""
End If
'Backbuffer.DrawText 50, 30, msgcount, False
DoEvents
If Ok% = -1 Then GoTo Fin
Primary.Flip Nothing, DDFLIP_WAIT

'GAMMA
  If GammaSupport = True Then
  UpdateGamma Int(CurrRed), Int(CurrGreen), Int(CurrBlue)
  End If
If JEU.statut = "menuprincipal" Then CurrRed = 0: CurrBlue = 0: CurrGreen = 0
If JEU.statut = "" And CurrRed < 0 Then
CurrRed = CurrRed + 0.5
CurrBlue = CurrBlue + 0.5
CurrGreen = CurrGreen + 0.5
End If
If JEU.statut = "gameover" Then
If CurrRed > -100 And CurrBlue < -50 Then CurrRed = CurrRed - 1
If CurrBlue > -100 Then CurrBlue = CurrBlue - 2: CurrGreen = CurrGreen - 2
End If

'si le perso est sur les bords de l'écran, camlock=true
camlock_H = True
camlock_B = True
camlock_R = True
camlock_L = True
If persoX < 300 And (perso_dir = "L" Or perso_dir = "HL" Or perso_dir = "BL") And PosMondeX * -1 / 32 >= 0 Then camlock_L = False
If persoX > 500 And (perso_dir = "R" Or perso_dir = "HR" Or perso_dir = "BR") And PosMondeX * -1 / 32 - longueurMapX + screenwidth / 32 < -1 Then camlock_L = False
If persoY < 250 And (perso_dir = "H" Or perso_dir = "HL" Or perso_dir = "HR") And PosMondeY * -1 / 32 >= 0 Then camlock_H = False
If persoY > 350 And (perso_dir = "B" Or perso_dir = "BL" Or perso_dir = "BR") And PosMondeY * -1 / 32 - longueurMapY + screenheight / 32 < 0 Then camlock_H = False


If longueurMapX * 32 < screenwidth Or longueurMapY * 32 < screenheight Then
PosMondeX = screenwidth - (longueurMapX + 4) * 32
PosMondeY = screenheight - (longueurMapY + 5) * 32
End If

If perf.IsPlaying(seg, segstate) = False Then Call perf.PlaySegment(seg, 0, 0)           ' On lance la lecture du fichier MIDI

Loop
Fin:
Unloade: Unload Form1
End Sub

Sub Unloade()
  DD.RestoreDisplayMode

 Set Primary = Nothing
 Set Backbuffer = Nothing
 
 Set DX = Nothing
 Set DD = Nothing
 Set DS = Nothing
End Sub
Public Function ddRect(x1, y1, x2, y2) As RECT
With ddRect
.Left = x1: .Right = x2: .Top = y1: .Bottom = y2
End With
End Function
Public Sub LoadJEU()
Set DD = DX.DirectDrawCreate("")
DD.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT
DD.SetDisplayMode screenwidth, screenheight, 32, 0, DDSDM_DEFAULT

Dim ddsd As DDSURFACEDESC2
ddsd.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
ddsd.lBackBufferCount = 1
ddsd.ddscaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE Or DDSCAPS_VIDEOMEMORY
Set Primary = DD.CreateSurface(ddsd)
Dim ddscaps As DDSCAPS2
ddscaps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_VIDEOMEMORY
Set Backbuffer = Primary.GetAttachedSurface(ddscaps)
Form1.Show
'----------------------------------------------------------------------------
'Là, c'est la partie Chargement...
'ON CHARGE TOUTES LES SURFACES
'----------------------------------------------        TILES
Set Tile(0) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile0.bmp", Tileddsd)
Set Tile(1) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile1.bmp", Tileddsd)
Set Tile(2) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile2.bmp", Tileddsd)
Set Tile(3) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile3.bmp", Tileddsd)
Set Tile(4) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile4.bmp", Tileddsd)
Set Tile(5) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile5.bmp", Tileddsd)
Set Tile(6) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile6.bmp", Tileddsd)
Set Tile(7) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile7.bmp", Tileddsd)
Set Tile(8) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile8.bmp", Tileddsd)
Set Tile(9) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile9.bmp", Tileddsd)
Set Tile(10) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile10.bmp", Tileddsd)
Set Tile(11) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile11.bmp", Tileddsd)
Set Tile(12) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile12.bmp", Tileddsd)
Set Tile(13) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile13.bmp", Tileddsd)
Set Tile(14) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile14.bmp", Tileddsd)
Set Tile(17) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile17.bmp", Tileddsd)
Set Tile(18) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile18.bmp", Tileddsd)
Set Tile(19) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile19.bmp", Tileddsd)
Set Tile(20) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile20.bmp", Tileddsd)
Set Tile(21) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile21.bmp", Tileddsd)
Set Tile(22) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile22.bmp", Tileddsd)
Set Tile(23) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile23.bmp", Tileddsd)
Set Tile(24) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile24.bmp", Tileddsd)

Set Tile(101) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile101.bmp", Tileddsd)
Set Tile(102) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile102.bmp", Tileddsd)
Set Tile(103) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile103.bmp", Tileddsd)
Set Tile(104) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile104.bmp", Tileddsd)
Set Tile(105) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile105.bmp", Tileddsd)
Set Tile(106) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile106.bmp", Tileddsd)
Set Tile(107) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile107.bmp", Tileddsd)
Set Tile(108) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile108.bmp", Tileddsd)
Set Tile(109) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile109.bmp", Tileddsd)
Set Tile(110) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile110.bmp", Tileddsd)
Set Tile(111) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile111.bmp", Tileddsd)
Set Tile(112) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile112.bmp", Tileddsd)
Set Tile(113) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile113.bmp", Tileddsd)
Set Tile(114) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile114.bmp", Tileddsd)
Set Tile(115) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile115.bmp", Tileddsd)
Set Tile(116) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile116.bmp", Tileddsd)
Set Tile(117) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile117.bmp", Tileddsd)
Set Tile(118) = DD.CreateSurfaceFromFile(App.Path & "\Map\Tile118.bmp", Tileddsd)

'----------------------------------------------        PERSO
Set Perso(1) = DD.CreateSurfaceFromFile(App.Path & "\perso\B1.bmp", Persoddsd(1))
Set Perso(2) = DD.CreateSurfaceFromFile(App.Path & "\perso\B2.bmp", Persoddsd(2))
Set Perso(3) = DD.CreateSurfaceFromFile(App.Path & "\perso\B3.bmp", Persoddsd(3))
Set Perso(4) = DD.CreateSurfaceFromFile(App.Path & "\perso\B4.bmp", Persoddsd(4))

Set perso_sword_R(1) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R1.bmp", Perso_swordddsd_R(1))
Set perso_sword_R(2) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R2.bmp", Perso_swordddsd_R(2))
Set perso_sword_R(3) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R3.bmp", Perso_swordddsd_R(3))
Set perso_sword_R(4) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R4.bmp", Perso_swordddsd_R(4))
Set perso_sword_R(5) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R5.bmp", Perso_swordddsd_R(5))
Set perso_sword_R(6) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R6.bmp", Perso_swordddsd_R(6))
Set perso_sword_R(7) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R7.bmp", Perso_swordddsd_R(7))
Set perso_sword_R(8) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_R7.bmp", Perso_swordddsd_R(8))

Set perso_sword_L(1) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L1.bmp", Perso_swordddsd_L(1))
Set perso_sword_L(2) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L2.bmp", Perso_swordddsd_L(2))
Set perso_sword_L(3) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L3.bmp", Perso_swordddsd_L(3))
Set perso_sword_L(4) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L4.bmp", Perso_swordddsd_L(4))
Set perso_sword_L(5) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L5.bmp", Perso_swordddsd_L(5))
Set perso_sword_L(6) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L6.bmp", Perso_swordddsd_L(6))
Set perso_sword_L(7) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L7.bmp", Perso_swordddsd_L(7))
Set perso_sword_L(8) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_L7.bmp", Perso_swordddsd_L(8))

Set perso_sword_H(1) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h1.bmp", Perso_swordddsd_H(1))
Set perso_sword_H(2) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h2.bmp", Perso_swordddsd_H(2))
Set perso_sword_H(3) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h3.bmp", Perso_swordddsd_H(3))
Set perso_sword_H(4) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h4.bmp", Perso_swordddsd_H(4))
Set perso_sword_H(5) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h5.bmp", Perso_swordddsd_H(5))
Set perso_sword_H(6) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h6.bmp", Perso_swordddsd_H(6))
Set perso_sword_H(7) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h7.bmp", Perso_swordddsd_H(7))
Set perso_sword_H(8) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_h7.bmp", Perso_swordddsd_H(8))

Set perso_sword_B(1) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B1.bmp", Perso_swordddsd_B(1))
Set perso_sword_B(2) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B2.bmp", Perso_swordddsd_B(2))
Set perso_sword_B(3) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B3.bmp", Perso_swordddsd_B(3))
Set perso_sword_B(4) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B4.bmp", Perso_swordddsd_B(4))
Set perso_sword_B(5) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B5.bmp", Perso_swordddsd_B(5))
Set perso_sword_B(6) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B6.bmp", Perso_swordddsd_B(6))
Set perso_sword_B(7) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B7.bmp", Perso_swordddsd_B(7))
Set perso_sword_B(8) = DD.CreateSurfaceFromFile(App.Path & "\perso\sword\sword_B7.bmp", Perso_swordddsd_B(8))
Set bombsurf = DD.CreateSurfaceFromFile(App.Path & "\perso\bombes\posee.bmp", bombsurfddsd)
Set bombsurf2 = DD.CreateSurfaceFromFile(App.Path & "\perso\bombes\posee2.bmp", bombsurfddsd)

Set explosion_surf = DD.CreateSurfaceFromFile(App.Path & "\gfx\explosion1.bmp", explosion_surfddsd)
'----------------------------------------------        CHAR
Set charsurf(1) = DD.CreateSurfaceFromFile(App.Path & "\char\1.bmp", charsurfddsd(1))
Set charsurf(2) = DD.CreateSurfaceFromFile(App.Path & "\perso\B2.bmp", charsurfddsd(2))
Set charsurf(3) = DD.CreateSurfaceFromFile(App.Path & "\perso\B3.bmp", charsurfddsd(3))
Set charsurf(4) = DD.CreateSurfaceFromFile(App.Path & "\perso\B4.bmp", charsurfddsd(4))
'-----------------------------------------------------------
Set OBJsurf(1) = DD.CreateSurfaceFromFile(App.Path & "\Objets\Buisson1.bmp", OBJsurfddsd(1))
'-----------------------------------------------------------
Set ENNEMIsurf(1) = DD.CreateSurfaceFromFile(App.Path & "\ennemis\ennemi1.bmp", ENNEMIsurfddsd(1))
Set Titre = DD.CreateSurfaceFromFile(App.Path & "\Menus\Titre.bmp", Titreddsd)

'couleur de transparence
ColorKey.high = vbWhite
ColorKey.low = vbWhite

For I = 1 To 4
Perso(I).SetColorKey DDCKEY_SRCBLT, ColorKey
Next I

For I = 1 To 8
perso_sword_L(I).SetColorKey DDCKEY_SRCBLT, ColorKey
perso_sword_R(I).SetColorKey DDCKEY_SRCBLT, ColorKey
perso_sword_B(I).SetColorKey DDCKEY_SRCBLT, ColorKey
perso_sword_H(I).SetColorKey DDCKEY_SRCBLT, ColorKey
Next I

OBJsurf(1).SetColorKey DDCKEY_SRCBLT, ColorKey
ENNEMIsurf(1).SetColorKey DDCKEY_SRCBLT, ColorKey
bombsurf.SetColorKey DDCKEY_SRCBLT, ColorKey
bombsurf2.SetColorKey DDCKEY_SRCBLT, ColorKey

explosion_surf.SetColorKey DDCKEY_SRCBLT, ColorKey

For j = 1 To 4
charsurf(j).SetColorKey DDCKEY_SRCBLT, ColorKey
Next j
Titre.SetColorKey DDCKEY_SRCBLT, ColorKey

PosMondeX = Val(LireIni("init", "posmonde_x", App.Path & "\Map\Map.map.prm"))
PosMondeY = Val(LireIni("init", "posmonde_y", App.Path & "\Map\Map.map.prm"))

Heros.statut.vie = 100
Heros.statut.vitesse = 4
Heros.Armes.epee = 1
Heros.Armes.bombes.nb = 30

perso_dir = "B"
perso_index = 1
bomb_counter = 1

JEU.statut = ""
If LireIni("init", "persox", App.Path & "\Map\Map.map.prm") <> "" Then persoX = LireIni("init", "persox", App.Path & "\Map\Map.map.prm")
If LireIni("init", "persoy", App.Path & "\Map\Map.map.prm") <> "" Then persoY = LireIni("init", "persoy", App.Path & "\Map\Map.map.prm")


LoadMAP ("\Map\Map.map")

'GAMMA (on crée le "gammacontroler" si l'ordi le permet
    CheckForGammaSupport
    If GammaSupport = True Then
    CreateGamma
    End If
  CurrRed = -100
  CurrBlue = -100
  CurrGreen = -100
End Sub
Sub move()
    If perso_move = True Then
    
    If Form1.Keyb.UpKey = True Then
    If sword_state = 0 Then
    perso_dir = "H"
    perso_index = 2
    If tileON_topleft < 100 And tileON_topright < 100 And tileON_topleft <> 4 And tileON_topright <> 4 And tileON_topleft <> 2 And tileON_topright <> 2 And tileON_topleft <> 11 And tileON_topright <> 11 And tileON_topleft <> 12 And tileON_topright <> 12 Then
    If camlock_B = False Then PosMondeY = PosMondeY + Heros.statut.vitesse Else persoY = persoY - Heros.statut.vitesse
    End If
    End If
    
    '------- Vérification des collision avec objets & persos
    
    For I = 1 To nbrChar
    
    If Int((PosMondeX * -1) / 32 + persoX / 32) = char(I).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = char(I).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = char(I).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = char(I).y Then
    If camlock_B = False Then PosMondeY = PosMondeY - Heros.statut.vitesse Else persoY = persoY + Heros.statut.vitesse
    End If
    End If
    Next I
    
    For j = 1 To nbrOBJ
    
    If Int((PosMondeX * -1) / 32 + persoX / 32) = OBJ(j).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(j).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = OBJ(j).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = OBJ(j).y Then
    If camlock_B = False Then PosMondeY = PosMondeY - Heros.statut.vitesse Else persoY = persoY + Heros.statut.vitesse
    End If
    End If
    Next j
    
    For K = 1 To nbrENNEMI
    
    If Int((PosMondeX * -1) / 32 + persoX / 32) = ENNEMI(K).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = ENNEMI(K).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = ENNEMI(K).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = ENNEMI(K).y Then
    If camlock_B = False Then PosMondeY = PosMondeY - 30 Else persoY = persoY + 30
    Heros.statut.vie = Heros.statut.vie - 3
    
    End If
    End If
    Next K
    End If
    
    If Form1.Keyb.DownKey = True Then
    If sword_state = 0 Then
    perso_dir = "B"
    perso_index = 1
    If tileON_bottomright < 100 And tileON_bottomleft < 100 And tileON_bottomright <> 4 And tileON_bottomleft <> 4 And tileON_bottomright <> 2 And tileON_bottomleft <> 2 And tileON_bottomright <> 11 And tileON_bottomleft <> 11 And tileON_bottomright <> 12 And tileON_bottomleft <> 12 Then
    If camlock_H = False Then PosMondeY = PosMondeY - Heros.statut.vitesse Else persoY = persoY + Heros.statut.vitesse
    End If
    End If
    
    '------- Vérification des collision avec objets & persos
    
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32 + persoX / 32) = char(I).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = char(I).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = char(I).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = char(I).y Then
    If camlock_H = False Then PosMondeY = PosMondeY + Heros.statut.vitesse Else persoY = persoY - Heros.statut.vitesse
    End If
    End If
    Next I
    
    For j = 1 To nbrOBJ
    If Int((PosMondeX * -1) / 32 + persoX / 32) = OBJ(j).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(j).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = OBJ(j).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = OBJ(j).y Then
    If camlock_H = False Then PosMondeY = PosMondeY + Heros.statut.vitesse Else persoY = persoY - Heros.statut.vitesse
    End If
    End If
    Next j
    
    For K = 1 To nbrENNEMI
    If Int((PosMondeX * -1) / 32 + persoX / 32) = ENNEMI(K).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = ENNEMI(K).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = ENNEMI(K).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = ENNEMI(K).y Then
    If camlock_H = False Then PosMondeY = PosMondeY + 30 Else persoY = persoY - 30
    Heros.statut.vie = Heros.statut.vie - 3
    End If
    End If
    Next K
    End If
    
    If Form1.Keyb.RightKey Then
    If sword_state = 0 Then
    perso_dir = "R"
    perso_index = 4
    If tileON_bottomright < 100 And tileON_topright < 100 And tileON_bottomright <> 4 And tileON_topright <> 4 And tileON_bottomright <> 2 And tileON_topright <> 2 And tileON_bottomright <> 11 And tileON_topright <> 11 And tileON_bottomright <> 12 And tileON_topright <> 12 Then
    If camlock_L = False Then PosMondeX = PosMondeX - Heros.statut.vitesse Else persoX = persoX + Heros.statut.vitesse
    End If
    End If
    
    '------- Vérification des collision avec objets & persos
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32 + persoX / 32) = char(I).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = char(I).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = char(I).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = char(I).y Then
    If camlock_L = False Then PosMondeX = PosMondeX + Heros.statut.vitesse Else persoX = persoX - Heros.statut.vitesse
    End If
    End If
    Next I
    
    For j = 1 To nbrOBJ
    If Int((PosMondeX * -1) / 32 + persoX / 32) = OBJ(j).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(j).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = OBJ(j).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = OBJ(j).y Then
    If camlock_L = False Then PosMondeX = PosMondeX + Heros.statut.vitesse Else persoX = persoX - Heros.statut.vitesse
    End If
    End If
    Next j
    
    For K = 1 To nbrENNEMI
    If Int((PosMondeX * -1) / 32 + persoX / 32) = ENNEMI(K).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = ENNEMI(K).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = ENNEMI(K).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = ENNEMI(K).y Then
    If camlock_L = False Then PosMondeX = PosMondeX + 30 Else persoX = persoX - 30
    Heros.statut.vie = Heros.statut.vie - 3
    End If
    End If
    Next K
    
    End If
    
    If Form1.Keyb.LeftKey Then
    If sword_state = 0 Then
    perso_dir = "L"
    perso_index = 3
    
    If tileON_topleft < 100 And tileON_bottomleft < 100 And tileON_topleft <> 4 And _
        tileON_bottomleft <> 4 And tileON_topleft <> 2 And tileON_bottomleft <> 2 And _
        tileON_topleft <> 11 And tileON_bottomleft <> 11 And tileON_topleft <> 12 And _
        tileON_bottomleft <> 12 Then
    If camlock_R = False Then PosMondeX = PosMondeX + Heros.statut.vitesse Else persoX = persoX - Heros.statut.vitesse
    End If
    End If
    
    '------- Vérification des collision avec objets & persos
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32 + persoX / 32) = char(I).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = char(I).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = char(I).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = char(I).y Then
    If camlock_R = False Then PosMondeX = PosMondeX - Heros.statut.vitesse Else persoX = persoX + Heros.statut.vitesse
    End If
    End If
    Next I
    
    For j = 1 To nbrOBJ
    If Int((PosMondeX * -1) / 32 + persoX / 32) = OBJ(j).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(j).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = OBJ(j).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = OBJ(j).y Then
    If camlock_R = False Then PosMondeX = PosMondeX - Heros.statut.vitesse Else persoX = persoX + Heros.statut.vitesse
    End If
    End If
    Next j
    
    For K = 1 To nbrENNEMI
    If Int((PosMondeX * -1) / 32 + persoX / 32) = ENNEMI(K).x Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = ENNEMI(K).x Then
    If Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = ENNEMI(K).y Or Int((PosMondeY * -1) / 32 + persoY / 32 + 3) = ENNEMI(K).y Then
    If camlock_R = False Then PosMondeX = PosMondeX - 30 Else persoX = persoX + 30
    Heros.statut.vie = Heros.statut.vie - 3
    End If
    End If
    Next K
    
    End If
    End If
    
    If Form1.Keyb.UpKey And Form1.Keyb.LeftKey And sword_state = 0 Then perso_dir = "HL"
    If Form1.Keyb.UpKey And Form1.Keyb.RightKey And sword_state = 0 Then perso_dir = "HR"
    If Form1.Keyb.DownKey And Form1.Keyb.LeftKey And sword_state = 0 Then perso_dir = "BL"
    If Form1.Keyb.DownKey And Form1.Keyb.RightKey And sword_state = 0 Then perso_dir = "BR"
    
    If Form1.Keyb.ShiftKey = True Then control_sword
    If Form1.Keyb.ShiftKey = False And Form1.Sword_timer.Enabled = False Then sword_possible = True
    If Form1.Keyb.ControlKey = True Then control_bomb
    If Form1.Keyb.ControlKey = False And bomb_timer(5) = 0 Then bomb_possible = True
    

    '------------||_//-___--\ /--------------------------------------
    '------------||=\ -|__-- |---------------------------------------  ESPACE
    '------------|| \\-|__---|---------------------------------------
    If Form1.Keyb.SpaceKey Then
    If msgcount <= 0 Then
    If perso_index = 3 Then
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 3) = char(I).x + 3 And Int((PosMondeY * -1 + 1) / 32) + Int(persoY / 32 + 1) = char(I).y Or Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 3) = char(I).x + 3 And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 2) = char(I).y Then msg = char(I).Txt
    Next I
    End If
     
    If perso_index = 4 Then
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 3) = char(I).x And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 1) = char(I).y Or Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 3) = char(I).x And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 2) = char(I).y Then msg = char(I).Txt
    Next I
    End If
    
    If perso_index = 2 Then
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 2) = char(I).x And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 1) = char(I).y + 1 Or Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 2) = char(I).x + 1 And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 1) = char(I).y + 1 Then msg = char(I).Txt
    Next I
    End If
    
    If perso_index = 1 Then
    For I = 1 To nbrChar
    If Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 2) = char(I).x And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 1) = char(I).y - 2 Or Int((PosMondeX * -1) / 32) + Int(persoX / 32 + 2) = char(I).x + 1 And Int((PosMondeY * -1) / 32) + Int(persoY / 32 + 1) = char(I).y - 2 Then msg = char(I).Txt
    Next I
    End If
    End If
    Else: If msgcount > 0 And msg = "" Then msgcount = msgcount - 1
    End If
    '---------------------------------------------------------------------
End Sub

Sub move_collision()
    If tileON_topleft = 4 Or tileON_topleft = 2 Or tileON_topleft > 100 Then
   
    If camlock_H = False Then
     PosMondeY = PosMondeY - 0.5
    End If
        If camlock_L = False Then
     PosMondeX = PosMondeX - 0.5
    End If
    
    If camlock_H = True Then
     persoY = persoY + 0.75
    End If
    If camlock_L = True Then
     persoX = persoX + 0.75
    End If
    End If
   
    If tileON_bottomleft = 4 Or tileON_bottomleft = 2 Or tileON_bottomleft > 100 Then
    
    If camlock_B = False Then
    PosMondeY = PosMondeY + 0.5
    End If
    If camlock_L = False Then
    PosMondeX = PosMondeX - 0.5
    End If
    
    If camlock_B = True Then
    persoY = persoY - 0.75
    End If
    If camlock_L = True Then
    persoX = persoX + 0.75
    End If
    End If
   
    If tileON_bottomright = 4 Or tileON_bottomright = 2 Or tileON_bottomright > 100 Then
    If camlock_R = False Then
    PosMondeX = PosMondeX + 0.5
    End If
    If camlock_B = False Then
    PosMondeY = PosMondeY + 0.5
    End If
    
    If camlock_R = True Then
    persoX = persoX - 0.75
    End If
    If camlock_B = True Then
    persoY = persoY - 0.75
    End If
    End If
    If tileON_topright = 4 Or tileON_topright = 2 Or tileON_topright > 100 Then
    If camlock_R = False Then
    PosMondeX = PosMondeX + 0.5
    End If
    If camlock_H = False Then
    PosMondeY = PosMondeY - 0.5
    End If
    
    If camlock_R = True Then
    persoX = persoX - 0.75
    End If
    If camlock_H = True Then
    persoY = persoY + 0.75
    End If
    
    End If
End Sub

Sub Check_Collision()
    On Error GoTo e
    Backbuffer.SetForeColor vbYellow
    tileON_bottomright = Map(Int((PosMondeX * -1 - 1) / 32 + persoX / 32 + 2), Int((PosMondeY * -1 - 1) / 32 + persoY / 32 + 2))
    tileON_topleft = Map(Int((PosMondeX * -1 + 1) / 32 + persoX / 32 + 1), Int((PosMondeY * -1 + 1) / 32 + persoY / 32 + 1))
    tileON_bottomleft = Map(Int((PosMondeX * -1 + 1) / 32 + persoX / 32 + 1), Int((PosMondeY * -1 - 1) / 32 + persoY / 32 + 2))
    tileON_topright = Map(Int((PosMondeX * -1 - 1) / 32 + persoX / 32 + 2), Int((PosMondeY * -1 + 1) / 32 + persoY / 32 + 1))
    Exit Sub
e:
    tileON_bottomright = 0
    tileON_topleft = 0
    tileON_bottomleft = 0
    tileON_topright = 0
End Sub

Public Sub Check_Warp()
    For I = 1 To nbrWarp
    If Int((PosMondeX * -1 - 1) / 32 + persoX / 32 + 2) = WarpZone(I).x Or Int((PosMondeX * -1 + 1) / 32 + persoX / 32 + 1) = WarpZone(I).x Then
    If Int((PosMondeY * -1 - 1) / 32 + persoY / 32 + 2) = WarpZone(I).y Or Int((PosMondeY * -1 + 1) / 32 + persoY / 32 + 1) = WarpZone(I).y Then
    persoX = 320
    persoY = 224
    
    PosMondeX = WarpZone(I).destX
    PosMondeY = WarpZone(I).destY
    LoadMAP (WarpZone(I).DestMap)
    bomb_timer(bomb_counter) = 0
    End If
    End If
    Next I
End Sub

Public Function LoadMAP(chemin As String)
    Dim TextLine
    Dim PosMapX As Integer
    Dim PosMapY As Integer
    Dim fichierini As String
    fichierini = App.Path & chemin & ".prm"
    longueurMapX = 0
    longueurMapY = 0
    longueurMapX = Val(LireIni("Longueurs", "Map_X", fichierini))
    longueurMapY = Val(LireIni("Longueurs", "Map_Y", fichierini))
    ReDim Map(longueurMapX, longueurMapY)
    LoadDM (chemin)
    loadWarp (chemin)
    loadChar (chemin)
    loadOBJ (chemin)
    loadENNEMI (chemin)
    param_map.nom = LireIni("Général", "Nom", fichierini)
    'on efface les valeurs de collision, car sinon quand on recharge une carte il y a des bugs
    For I = 1 To longueurMapX
    For j = 1 To longueurMapY
    Map(I, j) = 0
    Next j
    Next I
    For I = 1 To 5
    bomb_timer(I) = 0
    bomb_posee(I).x = 0
    bomb_posee(I).y = 0
    Next I
    PosMapX = 1
    PosMapY = 0
    'Backbuffer.DrawText 200, 100, "Chargement " & param_map.nom, False
    Open App.Path & chemin For Input As #1  ' Ouvre le fichier.
    Line Input #1, TextLine  ' Lit la ligne dans la variable.
    '----Affiche 1 ligne
    While PosMapY < longueurMapY
    Map(PosMapX, PosMapY) = Mid(TextLine, PosMapX * 3 - 2, 3)
    PosMapX = PosMapX + 1
    If PosMapX >= longueurMapX Then
    Line Input #1, TextLine  ' Lit la ligne dans la variable.
    PosMapX = 1
    PosMapY = PosMapY + 1
    End If
    Wend
    Close #1   ' Ferme le fichier.
    'on charge les musiques

End Function

Public Function InitDM()
Set loader = Nothing
' Si une lecture est en cour le loader contient des informations, donc on le remet à 0

Set loader = DX.DirectMusicLoaderCreate()
' Nous disons que la variable loader et un système de DirectX, de type chargement
' de musique

Set perf = DX.DirectMusicPerformanceCreate()
' Notre variable d'information du fichier est initialisée de même que le loader, avec la différence que celle-ci est de type Performance.

Call perf.Init(Nothing, 0)
' On associe la variable perf a un fichier MIDI, ici Nothing (donc aucun), et à 0 pour le handle.

perf.SetPort -1, 80
' Un index ainsi qu'une définition des canaux utilisée doivent être donné.


Dim FileName
FileName = App.Path & "\musiques\intro2-z3.mid"
Set seg = loader.LoadSegment(FileName)

' permet le chargement automatique des instruments utilisés par le fichier MIDI
perf.SetMasterAutoDownload True

' On dit que le fichier MIDI est de forme standard et non pas pour un fichier prévu exprès pour DirectMusic
seg.SetStandardMidiFile

perf.SetMasterVolume 15 * 50                ' Définition du volume
Call perf.PlaySegment(seg, 0, 0)            ' On lance la lecture du fichier MIDI


End Function

Public Function loadWarp(chemin)
    'Chargement des zones de téléportations
    Dim currentwarp As Integer
    Dim fichierini As String
    fichierini = App.Path & chemin & ".prm"
    nbrWarp = LireIni("Warp_Général", "nbrWarp", fichierini)
    ReDim WarpZone(nbrWarp + 1)
    While currentwarp <= nbrWarp
    currentwarp = currentwarp + 1
    WarpZone(currentwarp).index = currentwarp
    WarpZone(currentwarp).x = LireIni("Warp_" & currentwarp, "X", fichierini)
    WarpZone(currentwarp).y = LireIni("Warp_" & currentwarp, "Y", fichierini)
    WarpZone(currentwarp).DestMap = LireIni("Warp_" & currentwarp, "DestMap", fichierini)
    WarpZone(currentwarp).destX = Val(LireIni("Warp_" & currentwarp, "X2", fichierini))
    WarpZone(currentwarp).destY = Val(LireIni("Warp_" & currentwarp, "Y2", fichierini))
    Wend
End Function

Public Function loadChar(chemin)
    'Chargement des personnages du jeu
    Dim currentchar As Integer
    Dim fichierini As String
    fichierini = App.Path & chemin & ".prm"
    nbrChar = Val(LireIni("Char_Général", "nbrChar", fichierini))
    ReDim char(nbrChar + 1)
    While currentchar <= nbrChar
    currentchar = currentchar + 1
    char(currentchar).index = currentchar
    char(currentchar).x = LireIni("Char_" & currentchar, "X", fichierini)
    char(currentchar).y = LireIni("Char_" & currentchar, "Y", fichierini)
    char(currentchar).img = LireIni("Char_" & currentchar, "img", fichierini)
    char(currentchar).Txt = LireIni("Char_" & currentchar, "txt", fichierini)
    Wend
End Function

Public Function loadOBJ(chemin As String)
    'Chargement des personnages du jeu
    Dim currentOBJ As Integer
    Dim fichierini As String
    fichierini = App.Path & chemin & ".prm"
    nbrOBJ = Val(LireIni("OBJ_Général", "nbrOBJ", fichierini))
    ReDim OBJ(nbrOBJ + 1)
    While currentOBJ <= nbrOBJ
    currentOBJ = currentOBJ + 1
    OBJ(currentOBJ).x = LireIni("OBJ_" & currentOBJ, "X", fichierini)
    OBJ(currentOBJ).y = LireIni("OBJ_" & currentOBJ, "Y", fichierini)
    OBJ(currentOBJ).type = LireIni("OBJ_" & currentOBJ, "type", fichierini)
    Wend
End Function

Public Function afficheMAP()
    Dim PosMapX As String
    Dim PosMapY As String
    PosMapX = 1
    PosMapY = 0
    '----Affiche 1 ligne
    While PosMapY < longueurMapY
    If Map(PosMapX, PosMapY) <> 21 Then AfficherTile Tile(Map(PosMapX, PosMapY)), Tileddsd, ((PosMapX - 1) * 32) + PosMondeX, (PosMapY * 32) + PosMondeY, ddRect(0, 0, 0, 0)
    If Map(PosMapX, PosMapY) = 21 Then AfficherTile Tile(Map(PosMapX, PosMapY) + animtile_index), Tileddsd, ((PosMapX - 1) * 32) + PosMondeX, (PosMapY * 32) + PosMondeY, ddRect(0, 0, 0, 0)
    PosMapX = PosMapX + 1
    If PosMapX >= longueurMapX Then
    PosMapX = 1
    PosMapY = PosMapY + 1
    End If
    Wend
    For I = 1 To nbrOBJ
    If Val(OBJ(I).x) <> -1 And OBJ(I).type = "" Then AfficherImage OBJsurf(1), OBJsurfddsd(1), Val(OBJ(I).x) * 32 + PosMondeX, (Val(OBJ(I).y) - 1) * 32 + PosMondeY, ddRect(0, 0, 0, 0)
    If Val(OBJ(I).x) <> -1 And OBJ(I).type = "jarre" Then AfficherImage OBJsurf(1), OBJsurfddsd(1), Val(OBJ(I).x) * 32 + PosMondeX, (Val(OBJ(I).y) - 1) * 32 + PosMondeY, ddRect(0, 0, 0, 0)
    Next I
    afficheBOMB
    afficheHERO
    afficheCHAR
    afficheENNEMI
'    drawDEBUG
End Function

Public Function afficheCHAR()
    For I = 1 To nbrChar
    AfficherImage charsurf(char(I).img), charsurfddsd(char(I).img), char(I).x * 32 + PosMondeX, (char(I).y - 1) * 32 + PosMondeY, ddRect(0, 0, 0, 0)
    'si le heros est devant le char, il est dessiné après lui
    If (PosMondeY * -1 - 1) / 32 + persoY / 32 > char(I).y - 1 And (PosMondeX * -1 + 1) / 32 + persoX / 32 - 1 < char(I).x And (PosMondeX * -1 - 1) / 32 + persoX / 32 + 1 > char(I).x Then afficheHERO
    Next I
End Function

Function AfficherTile(Image As DirectDrawSurface7, ddsd As DDSURFACEDESC2, x As Integer, y As Integer, Cadre As RECT)
    Dim CadreImage As RECT
    'Si le CadreAffichage=0 on prend l'écran comme cadre :
    If Cadre.Right = 0 And Cadre.Bottom = 0 Then
        Cadre.Right = screenwidth
        Cadre.Bottom = screenheight
    End If
    'Valeurs par défaut du cadre de l'image
    CadreImage.Right = ddsd.lWidth
    CadreImage.Bottom = ddsd.lHeight
    'Le cadre de l'image ne doit pas dépasser le cadre d'affichage
    If x < Cadre.Left Then
        If x + ddsd.lWidth > 0 Then CadreImage.Left = Cadre.Left - x Else Exit Function
    End If
    If x + ddsd.lWidth > Cadre.Right Then
        If x < Cadre.Right Then CadreImage.Right = Cadre.Right - x Else Exit Function
    End If
    If y < Cadre.Top Then
        If y + ddsd.lHeight > 0 Then CadreImage.Top = Cadre.Top - y Else Exit Function
    End If
    If y + ddsd.lHeight > Cadre.Bottom Then
        If y < Cadre.Bottom Then CadreImage.Bottom = Cadre.Bottom - y Else Exit Function
    End If
    'Dessiner l'image
    Backbuffer.BltFast Min(x, 0), Min(y, 0), Image, CadreImage, DDBLTFAST_NOCOLORKEY
End Function

Function AfficherImage(Image As DirectDrawSurface7, ddsd As DDSURFACEDESC2, x As Integer, y As Integer, Cadre As RECT)
    Dim CadreImage As RECT
    'Si le CadreAffichage=0 on prend l'écran comme cadre :
    If Cadre.Right = 0 And Cadre.Bottom = 0 Then
        Cadre.Right = screenwidth
        Cadre.Bottom = screenheight
    End If
    'Valeurs par défaut du cadre de l'image
    CadreImage.Right = ddsd.lWidth
    CadreImage.Bottom = ddsd.lHeight
    'Le cadre de l'image ne doit pas dépasser le cadre d'affichage
    If x < Cadre.Left Then
        If x + ddsd.lWidth > 0 Then CadreImage.Left = Cadre.Left - x Else Exit Function
    End If
    If x + ddsd.lWidth > Cadre.Right Then
        If x < Cadre.Right Then CadreImage.Right = Cadre.Right - x Else Exit Function
    End If
    If y < Cadre.Top Then
        If y + ddsd.lHeight > 0 Then CadreImage.Top = Cadre.Top - y Else Exit Function
    End If
    If y + ddsd.lHeight > Cadre.Bottom Then
        If y < Cadre.Bottom Then CadreImage.Bottom = Cadre.Bottom - y Else Exit Function
    End If
    'Dessiner l'image
    Backbuffer.BltFast Val(Min(x, 0)), Val(Min(y, 0)), Image, CadreImage, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_NOCOLORKEY
End Function

Function Min(ByVal Valeur As Integer, ByVal ValeurMin As Integer) As Integer
    If Valeur < ValeurMin Then Min = ValeurMin Else Min = Valeur
End Function
Sub CheckForGammaSupport()
    Dim Hard As DDCAPS, Soft As DDCAPS
    DD.GetCaps Hard, Soft
    If (Hard.lCaps2 And DDCAPS2_PRIMARYGAMMA) = 0 Then
    GammaSupport = False
    Else
    GammaSupport = True
    End If
End Sub
Sub CreateGamma()
    If GammaSupport = False Then Exit Sub
    If GammaSupport = True Then
    Set GammaControler = Primary.GetDirectDrawGammaControl
    GammaControler.GetGammaRamp DDSGR_DEFAULT, OriginalRamp
    End If
End Sub

Sub UpdateGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)
'I'm not sure who wrote this procedure; but I (Jack Hoxley) didn't.
'Full credit to whoever did...
On Error GoTo GamOut:
Dim I As Integer

If GammaSupport = True Then
'Alter the gamma ramp to the percent given by comparing to original state
'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
'gamma level being set back to the original levels. Anything ABOVE zero will
'fade towards FULL colour, anything below zero will fade towards NO colour
For I = 0 To 255
    If intRed < 0 Then GammaRamp.red(I) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.red(I)) * (100 - Abs(intRed)) / 100)
    If intRed = 0 Then GammaRamp.red(I) = OriginalRamp.red(I)
    If intRed > 0 Then GammaRamp.red(I) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.red(I))) * (100 - intRed) / 100))
    If intGreen < 0 Then GammaRamp.green(I) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.green(I)) * (100 - Abs(intGreen)) / 100)
    If intGreen = 0 Then GammaRamp.green(I) = OriginalRamp.green(I)
    If intGreen > 0 Then GammaRamp.green(I) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.green(I))) * (100 - intGreen) / 100))
    If intBlue < 0 Then GammaRamp.blue(I) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.blue(I)) * (100 - Abs(intBlue)) / 100)
    If intBlue = 0 Then GammaRamp.blue(I) = OriginalRamp.blue(I)
    If intBlue > 0 Then GammaRamp.blue(I) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.blue(I))) * (100 - intBlue) / 100))
Next
GammaControler.SetGammaRamp DDSGR_DEFAULT, GammaRamp
End If
Exit Sub
GamOut:
End Sub
Private Function ConvToSignedValue(lngValue As Long) As Integer
'This was written by the same person who did the "updateGamma" code
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    ConvToSignedValue = CInt(lngValue - 65535)
End Function
Private Function ConvToUnSignedValue(intValue As Integer) As Long
'This was written by the same person who did the "updateGamma" code
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    ConvToUnSignedValue = intValue + 65535
End Function

Public Function drawDEBUG()
    'S'il vous plait votez pour moi en 2007!
    Backbuffer.DrawText 1, 1, "PosmondeX:" & PosMondeX, False
    Backbuffer.DrawText 1, 15, "PosmondeY:" & PosMondeY, False
    Backbuffer.DrawText 1, 30, "PersoX:" & persoX, False
    Backbuffer.DrawText 1, 45, "PersoY:" & persoY, False
    Backbuffer.DrawText 1, 75, "Vie:" & Heros.statut.vie, False
    Backbuffer.DrawText 1, 100, "Nombres de bombes:" & Heros.Armes.bombes.nb, False
    Backbuffer.DrawText 1, 125, "camlock:" & camlock, False
    Backbuffer.DrawText 1, 150, "tileON_topright:" & tileON_topright, False
    Backbuffer.DrawText 1, 175, "tileON_topleft:" & tileON_topleft, False
    Backbuffer.DrawText 1, 200, "tileON_bottomright:" & tileON_bottomright, False
    Backbuffer.DrawText 1, 225, "tileON_bottomleft:" & tileON_bottomleft, False
End Function
Public Sub LoadDS()
'cette sub n'est pas de moi!
'On Error GoTo Erreur ' ! important

'création de l'objet directSound
Set DS = DX.DirectSoundCreate("")
'priorité de l'application
DS.SetCooperativeLevel Form1.hWnd, DSSCL_PRIORITY

'format Wave : mono, 16 bits
With DSWaveFormat
    .nFormatTag = WAVE_FORMAT_PCM
    .nChannels = 2
    .lSamplesPerSec = 22050
    .nBitsPerSample = 16
    .nBlockAlign = .nBitsPerSample / 8 * .nChannels
    .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
End With

              ReDim BuffSons(4)
              ReDim DescSons(4)
DescSons(1).lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
DescSons(2).lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
DescSons(3).lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
'création du buffer
Set BuffSons(1) = DS.CreateSoundBufferFromFile(App.Path & "\sound\sword.wav", DescSons(1), DSWaveFormat)
Set BuffSons(2) = DS.CreateSoundBufferFromFile(App.Path & "\sound\dead.wav", DescSons(2), DSWaveFormat)
Set BuffSons(3) = DS.CreateSoundBufferFromFile(App.Path & "\sound\bomb.wav", DescSons(3), DSWaveFormat)

Exit Sub

Erreur:
MsgBox "Erreur d'initialisation de DirectSound." & vbCrLf & _
       "Aucune carte son détectée.", vbCritical, "Erreur !"
End

End Sub

Public Sub LoadDM(chemin As String)
If LireIni("Son", "FileName", App.Path & chemin & ".prm") <> "" Then
If DM_FileName <> App.Path & LireIni("Son", "FileName", App.Path & chemin & ".prm") Then
DM_FileName = App.Path & LireIni("Son", "FileName", App.Path & chemin & ".prm")
Set seg = loader.LoadSegment(DM_FileName)

' permet le chargement automatique des instruments utilisés par le fichier MIDI
perf.SetMasterAutoDownload True

' On dit que le fichier MIDI est de forme standard et non pas pour un fichier prévu exprès pour DirectMusic
seg.SetStandardMidiFile

Call perf.PlaySegment(seg, 0, 0)            ' On lance la lecture du fichier MIDI
End If
End If
End Sub

'si ça vous a plu, ecrivez a gwenael.pluchon@wanadoo.fr. J'accepte les cheques de 10 ou 20 euros,ainsi que les dons en nature.  Contactez moi d'abord par email. ;-)
'J'aimerais dire a tous mes fans d'arreter de m'attendre devant chez moi, et de stopper de remplir ma boite aux lettres. Merci

