Attribute VB_Name = "Mod_PERSO"
'---------------------------------------------------------------------------------------
' Module    : Mod_PERSO
' DateTime  : 30/12/2004 09:44
' Author    : Gwenael
'---------------------------------------------------------------------------------------

Public Sub control_sword()
'------------||_//-___--\ /--------------------------------------
'------------||=\ -|__-- |---------------------------------------  SHIFT
'------------|| \\-|__---|---------------------------------------
If Heros.Armes.epee <> 0 Then
If sword_possible = True Then
sword_possible = False
If perso_index = 3 Then
For I = 1 To nbrOBJ
If OBJ(I).type = "" And Int((PosMondeX * -1) / 32 + persoX / 32 + 2) = OBJ(I).x + 3 And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y - 1 Or Int((PosMondeX * -1) / 32 + persoX / 32 + 2) = OBJ(I).x + 3 And Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = OBJ(I).y - 1 Then OBJ(I).x = -1
Next I
End If

If perso_index = 4 Then
For I = 1 To nbrOBJ
If OBJ(I).type = "" And Int((PosMondeX * -1) / 32 + persoX / 32 + 2) = OBJ(I).x And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y - 1 Or Int((PosMondeX * -1) / 32 + persoX / 32 + 2) = OBJ(I).x And Int((PosMondeY * -1) / 32 + persoY / 32 + 2) = OBJ(I).y - 1 Then OBJ(I).x = -1
Next I
End If

If perso_index = 2 Then

For I = 1 To nbrOBJ
If OBJ(I).type = "" And Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(I).x And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(I).x + 1 And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(I).x And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y + 2 Then OBJ(I).x = -1
Next I
End If

If perso_index = 1 Then
For I = 1 To nbrOBJ
If OBJ(I).type = "" And Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(I).x And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y - 3 Or Int((PosMondeX * -1) / 32 + persoX / 32 + 1) = OBJ(I).x + 1 And Int((PosMondeY * -1) / 32 + persoY / 32 + 1) = OBJ(I).y - 3 Then OBJ(I).x = -1
Next I
End If
anim_sword = 1
sword_state = 1
Form1.Sword_timer.Enabled = True

BuffSons(1).Play (DSBPLAY_DEFAULT)
End If
End If
End Sub
Public Sub control_bomb()
'------------||_//-___--\ /--------------------------------------
'------------||=\ -|__-- |---------------------------------------  CTRL
'------------|| \\-|__---|---------------------------------------
If Form1.bomb_reload_timer.Enabled = False Then
If Heros.Armes.bombes.nb > 0 Then
If bomb_possible = True Then
If bomb_counter > 4 Then bomb_counter = 1

Form1.bomb_reload_timer.Enabled = True

bomb_timer(bomb_counter) = 1
bomb_possible = False
Form1.bomb_reload_timer.Enabled = False: Form1.bomb_reload_timer.Enabled = True
Heros.Armes.bombes.nb = Heros.Armes.bombes.nb - 1
bomb_posee(bomb_counter).x = (PosMondeX * -1 + persoX) / 32
bomb_posee(bomb_counter).y = (PosMondeY * -1 + persoY) / 32 + 1
bomb_timer(bomb_counter) = 1

If bomb_timer(bomb_counter) > 0 Then bomb_timer(bomb_counter) = bomb_timer(bomb_counter) + 1

bomb_counter = bomb_counter + 1
End If
End If
End If
End Sub

Public Sub afficheBOMB()
For j = 1 To 5

If bomb_timer(j) <> 0 Then
If bomb_timer(j) / 10 = Round(bomb_timer(j) / 10) Then AfficherImage bombsurf2, bombsurfddsd, bomb_posee(j).x * 32 + PosMondeX, bomb_posee(j).y * 32 + PosMondeY, ddRect(0, 0, 0, 0) Else AfficherImage bombsurf, bombsurfddsd, bomb_posee(j).x * 32 + PosMondeX, bomb_posee(j).y * 32 + PosMondeY, ddRect(0, 0, 0, 0)
bomb_timer(j) = bomb_timer(j) + 1
If bomb_timer(j) = 100 Then
' Code de l'explosion de la bombe
bomb_explosion_timer(j) = 1
bomb_timer(j) = 0
BuffSons(3).Play (DSBPLAY_DEFAULT)

For I = 1 To nbrOBJ

If OBJ(I).type = "" And Round(bomb_posee(j).x - 1) = OBJ(I).x Or Round(bomb_posee(j).x) = OBJ(I).x Or Round(bomb_posee(j).x + 1) = OBJ(I).x Then
If Round(bomb_posee(j).y + 1) = OBJ(I).y Or Round(bomb_posee(j).y) = OBJ(I).y Or Round(bomb_posee(j).y + 2) = OBJ(I).y Then OBJ(I).x = -1
End If
Next I

End If
End If

If bomb_explosion_timer(j) > 0 Then bomb_explosion_timer(j) = bomb_explosion_timer(j) + 1: AfficherImage explosion_surf, explosion_surfddsd, (bomb_posee(j).x - 1) * 32 + PosMondeX, (bomb_posee(j).y - 1) * 32 + PosMondeY, ddRect(0, 0, 0, 0)
If bomb_explosion_timer(j) > 30 Then bomb_explosion_timer(j) = 0

Next j
Backbuffer.DrawText 300, 1, bomb_counter, False
End Sub

Public Sub afficheHERO()
If sword_state = 0 Then AfficherImage Perso(perso_index), Persoddsd(perso_index), persoX, persoY, ddRect(0, 0, 0, 0): anim_sword = 0

If sword_state = 1 Then
If perso_dir = "B" Then AfficherImage perso_sword_B(Int(anim_sword)), Perso_swordddsd_B(anim_sword), persoX - 32, persoY + 16, ddRect(0, 0, 0, 0)
If perso_dir = "H" Then AfficherImage perso_sword_H(Int(anim_sword)), Perso_swordddsd_H(anim_sword), persoX - 32, persoY - 16, ddRect(0, 0, 0, 0)
If perso_dir = "L" Or perso_dir = "HL" Or perso_dir = "BL" Then AfficherImage perso_sword_L(Int(anim_sword)), Perso_swordddsd_L(anim_sword), persoX - 42, persoY, ddRect(0, 0, 0, 0)
If perso_dir = "R" Or perso_dir = "BR" Or perso_dir = "HR" Then AfficherImage perso_sword_R(Int(anim_sword)), Perso_swordddsd_R(anim_sword), persoX - 15, persoY + 5, ddRect(0, 0, 0, 0)
If anim_sword <= 7 Then anim_sword = anim_sword + 0.5 Else anim_sword = 1: sword_state = 0:    Form1.Sword_timer.Enabled = False

End If

End Sub
