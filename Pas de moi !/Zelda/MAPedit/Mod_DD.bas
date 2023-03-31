Attribute VB_Name = "Mod_DD"
'---------------------------------------------------------------------------------------
' Module    : Mod_DD
' DateTime  : 16/06/2004 11:44
' Author    : Gwenael
'
'Merci a AbeLeMudokon pour sa fonction AfficherImage très utile(ça m'a évité de me
'prendre la tete sur cette partie de code plutot chiante ;-))
'
'Je tiens a préciser que toutes les fonction de GAMMA ne sont pas de moi, mais du site
'http://www.dx4vb.da.ru . Merci a eux pour leur tutorial!
'
'Pour le moment il n'y a des collision que sur un types de tile d'eau (la N°4). On peut rajouter
'd'autres surfaces facilement mais je n'en ai mis qu'une pour tester les collisions +
'facilement.
'
'DESOLE POUR LE MANQUE DE COMMENTAIRES (LA FLEMME ;-) )
'
'SVP LAISSEZ DES COMMENTAIRES SUR VBFRANCE CE SERAIT SYMPA(ET CA PERMET DE FAIRE
'PROGRESSER LA SCOURCE!)
'
'---------------------------------------------------------------------------------------
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public DX As New DirectX7
Public DD As DirectDraw7
Public Primary As DirectDrawSurface7    'Surface primaire visible a l'écran
Public Backbuffer As DirectDrawSurface7 'Surface brouillon invisible
Public Ok As Integer
Public ColorKey As DDCOLORKEY

Public Tile(10) As DirectDrawSurface7
Public Tileddsd As DDSURFACEDESC2
Public Perso As DirectDrawSurface7
Public cur_alpha As DirectDrawSurface7
Public cur_alphaddsd As DDSURFACEDESC2
Public Persoddsd As DDSURFACEDESC2

Public menu1 As DirectDrawSurface7
Public menu1ddsd As DDSURFACEDESC2

Public persoX As Integer
Public persoY As Integer

Public PosMondeX
Public PosMondeY

Public tileON_topleft
Public tileON_X
Public tileON_Y

Public StopJeu As Boolean

Public screenwidth
Public screenheight

Public Map(150, 150)

Public Type param_map
 nom As String
 pathsubtile As String
 pathsurtile As String
End Type

Public param_map As param_map

Sub Main()
Dim fichierini As String

fichierini = App.Path & "\map\map.prm"

param_map.nom = LireIni("Général", "Nom", fichierini)
param_map.pathsubtile = LireIni("Général", "pathsubtile", fichierini)

screenwidth = 640
screenheight = 480

PosMondeY = 42

  persoX = 320
  persoY = 224
  
LoadJEU

ShowCursor 0
Backbuffer.SetFont Form1.Font


Do

Backbuffer.BltColorFill ddRect(0, 0, 0, 0), 0
Backbuffer.SetForeColor vbYellow

afficheMAP

Backbuffer.DrawText 50, 50, tileON, False

'------------------------------------------------
move
'------------------------------------------------
curseur_collision

DisplayFx cur_alpha, persoX, persoY, 22, 48, vbMergePaint, False
AfficherImage Perso, Persoddsd, persoX, persoY, ddRect(0, 0, 0, 0)
AfficherImage menu1, menu1ddsd, 10, 435, ddRect(0, 0, 0, 0)
DoEvents
If Ok% = -1 Then GoTo Fin

afficher_menu_tiles

Primary.Flip Nothing, DDFLIP_WAIT

Loop

Fin:

Unloade
Unload Form1
End Sub

Sub Unloade()
ShowCursor 1
 Set DD = Nothing
 Set Primary = Nothing
 Set Backbuffer = Nothing
 Set DX = Nothing
End Sub

Public Function ddRect(x1, y1, x2, y2) As RECT
With ddRect
.Left = x1
.Right = x2
.Top = y1
.Bottom = y2
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

Set Perso = DD.CreateSurfaceFromFile(App.Path & "\perso\curseur.bmp", Persoddsd)
Set cur_alpha = DD.CreateSurfaceFromFile(App.Path & "\perso\cur_ALPHA.bmp", cur_alphaddsd)

Set menu1 = DD.CreateSurfaceFromFile(App.Path & "\system\menu1.bmp", menu1ddsd)

'couleur de transparence
ColorKey.high = RGB(255, 255, 255)
ColorKey.low = RGB(255, 255, 255)

Perso.SetColorKey DDCKEY_SRCBLT, ColorKey
cur_alpha.SetColorKey DDCKEY_SRCBLT, ColorKey

menu1.SetColorKey DDCKEY_SRCBLT, ColorKey

LoadMAP (param_map.pathsubtile)
End Sub

Sub move()
On Error GoTo e

If Form1.Keyb.UpKey And PosMondeY < 42 Then PosMondeY = PosMondeY + 4.5

If Form1.Keyb.DownKey And PosMondeY > longueurMapY Then PosMondeY = PosMondeY - 4.5

If Form1.Keyb.RightKey And PosMondeX > longueurMapX Then
PosMondeX = PosMondeX - 4.5
End If

If Form1.Keyb.LeftKey And PosMondeX < 0 Then
PosMondeX = PosMondeX + 4.5
End If


If Form1.Keyb.num1Key Then Map(tileON_X + 1, tileON_Y) = "001"
If Form1.Keyb.num2Key Then Map(tileON_X + 1, tileON_Y) = "002"
If Form1.Keyb.num3Key Then Map(tileON_X + 1, tileON_Y) = "003"
If Form1.Keyb.num4Key Then Map(tileON_X + 1, tileON_Y) = "004"
If Form1.Keyb.num5Key Then Map(tileON_X + 1, tileON_Y) = "005"
If Form1.Keyb.num6Key Then Map(tileON_X + 1, tileON_Y) = "006"
If Form1.Keyb.num7Key Then Map(tileON_X + 1, tileON_Y) = "007"
If Form1.Keyb.num8Key Then Map(tileON_X + 1, tileON_Y) = "008"
If Form1.Keyb.num9Key Then Map(tileON_X + 1, tileON_Y) = "009"
If Form1.Keyb.num0Key Then Map(tileON_X + 1, tileON_Y) = "010"

If persoX > 638 And PosMondeX > longueurMapX Then PosMondeX = PosMondeX - 8
If persoX < 2 And PosMondeX < 0 Then PosMondeX = PosMondeX + 8

If persoY > 478 And PosMondeY > longueurMapY - 42 Then PosMondeY = PosMondeY - 8
If persoY < 2 And PosMondeY < 42 Then PosMondeY = PosMondeY + 8

Exit Sub
e:

End Sub
Sub curseur_collision()
On Error GoTo e
tileON_X = Int((PosMondeX * -1 + 1) / 32) + Int(persoX) / 32
tileON_Y = Int((PosMondeY * -1 - 10) / 32) + Int(persoY) / 32
tileON_topleft = Map(Int((PosMondeX * -1 + 1) / 32) + 10, Int((PosMondeY * -1 + 1) / 32) + 7)

Exit Sub
e:
Backbuffer.DrawText 1, 20, "Pas de tile sous le curseur.", False
End Sub
Private Function LoadMAP(chemin As String)
On Error GoTo e
Dim TextLine

Dim PosMapX
Dim PosMapY
Dim longueurMapX
Dim longueurMapY


longueurMapX = 17
longueurMapY = 13

PosMapX = 1
PosMapY = 0
Backbuffer.DrawText 200, 100, "Chargement :" & param_map.nom, False
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
Exit Function
e:
Close #1
End Function

Private Function afficheMAP()
On Error GoTo e
Dim TextLine
Dim IndexSurface
Dim PosMapX
Dim PosMapY
Dim longueurMapX
Dim longueurMapY

longueurMapX = 17
longueurMapY = 13

PosMapX = 1
PosMapY = 0

'----Affiche 1 ligne
While PosMapY < longueurMapY

AfficherTile Tile(Map(PosMapX, PosMapY)), Tileddsd, ((PosMapX - 1) * 32) + PosMondeX, (PosMapY * 32) + PosMondeY, ddRect(0, 0, 0, 0)


PosMapX = PosMapX + 1
If PosMapX >= longueurMapX Then
PosMapX = 1
PosMapY = PosMapY + 1
End If
Wend
Exit Function
e:
'MsgBox "affichemap"
End Function

Function Min(ByVal Valeur As Integer, ByVal ValeurMin As Integer) As Integer

If Valeur < ValeurMin Then Min = ValeurMin Else Min = Valeur

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

Backbuffer.BltFast Min(x, 0), Min(y, 0), Image, CadreImage, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_NOCOLORKEY

End Function
