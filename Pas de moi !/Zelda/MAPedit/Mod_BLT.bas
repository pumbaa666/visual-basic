Attribute VB_Name = "Mod_BLT"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'----------------------------------------------------------------------------------
'----------------------------------------BLTFX-------------------------------------
Public Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Integer, intY As Integer, intWidth As Integer, intHeight As Integer, lngROP As Long, blnFxCap As Boolean)
Dim mrectScreen As RECT                     'Rectangle the size of the screen
Dim rectSource As RECT
Dim rectDest As RECT
Dim lngSrcDC As Long
Dim lngDestDC As Long
Dim objBltFx As DDBLTFX

    With rectSource
        'Set and clip the source rect
        If intY < 0 Then .Top = mrectScreen.Top - intY
        .Bottom = intHeight
        If intY + intHeight > mrectScreen.Bottom Then .Bottom = intHeight - ((intY + intHeight) - mrectScreen.Bottom)
        If intX < 0 Then .Left = mrectScreen.Left - intX
        .Right = intWidth
        If intX + intWidth > mrectScreen.Right Then .Right = intWidth - ((intX + intWidth) - mrectScreen.Right)
    End With

    With rectDest
        'Set and clip the destination rect
        .Top = intY
        If .Top < mrectScreen.Top Then .Top = mrectScreen.Top
        .Bottom = intY + intHeight
        If .Bottom > mrectScreen.Bottom Then .Bottom = mrectScreen.Bottom
        .Left = intX
        If .Left < mrectScreen.Left Then .Left = mrectScreen.Left
        .Right = intX + intWidth
        If .Right > mrectScreen.Right Then .Right = mrectScreen.Right
    End With

    'Can we use hardware acceleration?
    If blnFxCap Then
        objBltFx.lROP = lngROP
        Backbuffer.BltFx rectDest, surfDisplay, rectSource, DDBLT_ROP Or DDBLT_WAIT, objBltFx
    Else
        'If no hardware accel, do it the hard way: First, Get our DCs
        lngDestDC = Backbuffer.GetDC
        lngSrcDC = surfDisplay.GetDC
        'Do the fancy old-fashioned blit
        BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, 0, 0, lngROP
        'Release our DCs
        surfDisplay.ReleaseDC lngSrcDC
        Backbuffer.ReleaseDC lngDestDC
    End If

End Sub

Public Sub afficher_menu_tiles()
'MENU DES TILES
Backbuffer.BltColorFill ddRect(0, 0, 640, 42), RGB(75, 25, 0)

DisplayFx Tile(1 + posmenu_tile * 10), 130, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(2 + posmenu_tile * 10), 164, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(3 + posmenu_tile * 10), 198, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(4 + posmenu_tile * 10), 232, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(5 + posmenu_tile * 10), 266, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(6 + posmenu_tile * 10), 300, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(7 + posmenu_tile * 10), 334, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(8 + posmenu_tile * 10), 368, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(9 + posmenu_tile * 10), 402, 5, 32, 32, vbSrcCopy, False
DisplayFx Tile(10 + posmenu_tile * 10), 436, 5, 32, 32, vbSrcCopy, False

Backbuffer.SetForeColor 0

Backbuffer.DrawText 129, 21, "1", False
Backbuffer.DrawText 163, 21, "2", False
Backbuffer.DrawText 197, 21, "3", False
Backbuffer.DrawText 231, 21, "4", False
Backbuffer.DrawText 265, 21, "5", False
Backbuffer.DrawText 299, 21, "6", False
Backbuffer.DrawText 333, 21, "7", False
Backbuffer.DrawText 367, 21, "8", False
Backbuffer.DrawText 401, 21, "9", False
Backbuffer.DrawText 435, 21, "10", False

Backbuffer.DrawText 131, 23, "1", False
Backbuffer.DrawText 165, 22, "2", False
Backbuffer.DrawText 199, 22, "3", False
Backbuffer.DrawText 233, 22, "4", False
Backbuffer.DrawText 267, 22, "5", False
Backbuffer.DrawText 301, 22, "6", False
Backbuffer.DrawText 335, 22, "7", False
Backbuffer.DrawText 369, 22, "8", False
Backbuffer.DrawText 403, 22, "9", False
Backbuffer.DrawText 437, 22, "10", False

Backbuffer.SetForeColor RGB(255, 255, 255)

Backbuffer.DrawText 130, 22, "1", False
Backbuffer.DrawText 164, 22, "2", False
Backbuffer.DrawText 198, 22, "3", False
Backbuffer.DrawText 232, 22, "4", False
Backbuffer.DrawText 266, 22, "5", False
Backbuffer.DrawText 300, 22, "6", False
Backbuffer.DrawText 334, 22, "7", False
Backbuffer.DrawText 368, 22, "8", False
Backbuffer.DrawText 402, 22, "9", False
Backbuffer.DrawText 436, 22, "10", False



End Sub
