Attribute VB_Name = "Mod_MAP"
Public Function map_SAVE()
On Error GoTo e
Dim TextLine

Dim PosMapX
Dim PosMapY
Dim longueurMapX
Dim longueurMapY

longueurMapX = 21
longueurMapY = 15

PosMapX = 0
PosMapY = 0

Backbuffer.DrawText 200, 100, "Sauvegarde :" & param_map.nom, False

Open App.Path & "\map.map" For Output As #1  ' Ouvre le fichier.
'       ------------
'       |  WRITE # |
'       ------------
Do Until PosMapY = longueurMapY


  While PosMapX <= longueurMapX
  TextLine = TextLine & Map(PosMapX, PosMapY)
  PosMapX = PosMapX + 1
  Wend

Print #1, TextLine
PosMapY = PosMapY + 1
PosMapX = 0
TextLine = ""
Loop
Close #1   ' Ferme le fichier.
Exit Function
e:
Unloade
MsgBox "Erreur" & Err.Description
End Function


