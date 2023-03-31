VERSION 5.00
Object = "{5C220880-255F-4A5E-B713-E9FD99E876B8}#1.0#0"; "PaintedBalls.ocx"
Object = "{8F5AA0BC-ED76-4B08-929A-26908E1EA235}#1.0#0"; "ScreenShoots.ocx"
Begin VB.Form frmMere 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMere"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ScreenShoot.ScreenShoots ScreenShoots1 
      Left            =   1920
      Top             =   3840
      _ExtentX        =   794
      _ExtentY        =   476
   End
   Begin PaintedBall.PaintedBalls PaintedBalls1 
      Left            =   1680
      Top             =   1800
      _ExtentX        =   1323
      _ExtentY        =   1323
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   5040
   End
End
Attribute VB_Name = "frmMere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'————————————————————————————————————————————————————————————————
'————————————————————————————————————————————————————————————————
'————————————————————————————————————————————————————————————————
'— Par ScSami dit BgBS dit A-Cube : Apatride, Asocial, Anarchiste
'— scsami@yahoo.fr
'— FREEWARE 2005
'————————————————————————————————————————————————————————————————
'————————————————————————————————————————————————————————————————
'————————————————————————————————————————————————————————————————
'— <ECHAP> pour quitter
'— Double Clique pour rafraîchir l'arrière plan
'— Ce programme démontre toute la puissance des Contrôles ActiveX
'— ScreenShoot et PaintedBalls que j'ai réalisés.
'— Le code est un peu lourd et il vaut mieux avoir une bonne
'— bécane.
'— Pour plus d'infos sur le code, m'écrire.
'————————————————————————————————————————————————————————————————
'————————————————————————————————————————————————————————————————
Private Const PI As Single = 3.14159265358979

Private varAutoChange As Boolean
Private varAutoChangeDblClick As Boolean
Private varTime As Single

Private tblStr(4) As String  'Matrice du Texte

Private varNbrPtsOrg As Byte
Private tblPtsOrg(149, 2) As Single

Private varNbrPtsAfi As Byte
Private tblPtsAfi(149, 3) As Single
Private tblPtsOrder(149) As Byte

Private varModeAffi As Byte
Private varColorBall As Byte
Private varTexture As Byte

'Centre de l'écran
Private varCX As Single, varCY As Single
'Coefs de profondeurs 2D
Private varCoZ As Single
'Décalage 3D sur plan
Private varX As Single
Private varY As Single
Private varZ As Single
'Centre du Décalage par Rotation 3D
Private varCRx As Single
Private varCRy As Single
Private varCRz As Single
'Angles de Décalage par Rotation 3D
Private varAngDegRotX As Single
Private varAngDegRotY As Single
Private varAngDegRotZ As Single
'Angles pour le contrôle clavier
Private varCoRx As Single
Private varCoRy As Single
Private varCoRz As Single

'Coef de la Taille Min. des "Méta-Balls"
Private varCoTxy As Single



Private Sub Form_DblClick()
 'Rafraîchir l'arrière plan
 If varAutoChangeDblClick = True Then Exit Sub
 varAutoChangeDblClick = True
 frmMere.Visible = False
 DoEvents
 ScreenShoots1.ShootScreen
 DoEvents
 frmMere.Visible = True
 frmMere.Picture = ScreenShoots1.PictureImage
 varAutoChangeDblClick = False
End Sub



Private Sub Form_Load()
 ScreenShoots1.ShootScreen
 DoEvents
 frmMere.Visible = True
 frmMere.Picture = ScreenShoots1.PictureImage
 Call Initialiser
 Timer1.Enabled = True
End Sub

Private Sub Initialiser()
 varTime = Timer
 varHeure = ""
 varAutoChange = False
 varAutoChangeDblClick = False
 
 varCX = frmMere.ScaleWidth \ 2
 varCY = frmMere.ScaleHeight \ 2
 
 varCoZ = 1.001
 varCoTxy = 50 / 3
 
 varX = 0
 varY = 0
 varZ = 500
 
 varCRx = 0
 varCRy = 0
 varCRz = 0
 'Vitesses de rotation
 varAngDegRotX = 0: varCoRx = 8
 varAngDegRotY = 0: varCoRy = 4
 varAngDegRotZ = 0: varCoRz = 2
 
 'Initialise le H et le ' dans le Tbl String constituant le texte heure
 For t = 0 To 4  '5 lignes
  tblStr(t) = String(30, " ")
 Next t
 'H
 Mid(tblStr(0), 9, 3) = ". ."
 Mid(tblStr(1), 9, 3) = ". ."
 Mid(tblStr(2), 9, 3) = "..."
 Mid(tblStr(3), 9, 3) = ". ."
 Mid(tblStr(4), 9, 3) = ". ."
 '"'"
 Mid(tblStr(0), 21, 2) = " ."
 Mid(tblStr(1), 21, 2) = " ."
 Mid(tblStr(2), 21, 2) = ". "
 
 varNbrPtsOrg = 0
 varNbrPtsAfi = 0
 
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyEscape Then End  'Fin du programme
End Sub



Private Sub Timer1_Timer()
 If varAutoChange = True Then Exit Sub
 varAutoChange = True
 
 'Si l'heure a changé, redéfinir le tableau de "Méta-Balls"
 If varTime <> Fix(Timer) Then Call InitTblTimer
 
 'Effectue une rotation automatique toutes les X ms
 varAngDegRotX = varAngDegRotX - varCoRx
 varAngDegRotY = varAngDegRotY + varCoRy
 varAngDegRotZ = varAngDegRotZ + varCoRz

 If varAngDegRotX < 0 Then varAngDegRotX = 360 + varAngDegRotX
 If varAngDegRotY < 0 Then varAngDegRotY = 360 + varAngDegRotY
 If varAngDegRotZ < 0 Then varAngDegRotZ = 360 + varAngDegRotZ

 If varAngDegRotX > 360 Then varAngDegRotX = varAngDegRotX - 360
 If varAngDegRotY > 360 Then varAngDegRotY = varAngDegRotY - 360
 If varAngDegRotZ > 360 Then varAngDegRotZ = varAngDegRotZ - 360
 
 Call Afficher
 
 varAutoChange = False
End Sub


Private Sub InitTblTimer()
 Static vH As Long
 Static vM As Long
 Static vS As Long
 Static vHstr1 As String, vHstr2 As String
 Static vMstr1 As String, vMstr2 As String
 Static vSstr1 As String, vSstr2 As String
 
 varTime = Fix(Timer)
 vH = varTime \ 3600
 vM = Fix((varTime - (vH * 3600)) / 60)
 vS = varTime - (vH * 3600) - (vM * 60)
 
 If vH < 9 Then
  vHstr1 = "0"
  vHstr2 = CStr(vH)
 Else
  vHstr1 = CStr(Fix(vH / 10))
  vHstr2 = CStr(vH - (Fix(vH / 10) * 10))
 End If
 Call InitChiffres(1, vHstr1)
 Call InitChiffres(5, vHstr2)
 
 If vM < 9 Then
  vMstr1 = "0"
  vMstr2 = CStr(vM)
 Else
  vMstr1 = CStr(Fix(vM / 10))
  vMstr2 = CStr(vM - (Fix(vM / 10) * 10))
 End If
 Call InitChiffres(13, vMstr1)
 Call InitChiffres(17, vMstr2)
 
 If vS < 9 Then
  vSstr1 = "0"
  vSstr2 = CStr(vS)
 Else
  vSstr1 = CStr(Fix(vS / 10))
  vSstr2 = CStr(vS - (Fix(vS / 10) * 10))
 End If
 Call InitChiffres(24, vSstr1)
 Call InitChiffres(28, vSstr2)
 
 varNbrPtsOrg = 0
 For t = 0 To 4
  For tt = 1 To 30
   If Mid(tblStr(t), tt, 1) <> " " Then
    tblPtsOrg(varNbrPtsOrg, 0) = (tt * 20) - 300 'X
    tblPtsOrg(varNbrPtsOrg, 1) = (-(t * 20)) + 50 'Y
    tblPtsOrg(varNbrPtsOrg, 2) = 0       'Z
    varNbrPtsOrg = varNbrPtsOrg + 1
   End If
  Next tt
 Next t
 varNbrPtsOrg = varNbrPtsOrg - 1
 
 If ((vS / 2) - Fix(vS / 2)) = 0 Then
  varModeAffi = Fix(Rnd * 2)
  If varModeAffi = 0 Then
   varColorBall = Fix(Rnd * 7)
   varTexture = Fix(Rnd * 16) + 1
  Else
   varTexture = Fix(Rnd * 35) + 1
  End If
 End If
End Sub


Private Sub InitChiffres(ByVal DebChar As Byte, ByVal CharNum As String)
 Select Case CharNum
 Case "0"
  Mid(tblStr(0), DebChar, 3) = " @ "
  Mid(tblStr(1), DebChar, 3) = "@ @"
  Mid(tblStr(2), DebChar, 3) = "@ @"
  Mid(tblStr(3), DebChar, 3) = "@ @"
  Mid(tblStr(4), DebChar, 3) = " @ "
 Case "1"
  Mid(tblStr(0), DebChar, 3) = " @ "
  Mid(tblStr(1), DebChar, 3) = "@@ "
  Mid(tblStr(2), DebChar, 3) = " @ "
  Mid(tblStr(3), DebChar, 3) = " @ "
  Mid(tblStr(4), DebChar, 3) = "@@@"
 Case "2"
  Mid(tblStr(0), DebChar, 3) = " @ "
  Mid(tblStr(1), DebChar, 3) = "@ @"
  Mid(tblStr(2), DebChar, 3) = "  @"
  Mid(tblStr(3), DebChar, 3) = " @ "
  Mid(tblStr(4), DebChar, 3) = "@@@"
 Case "3"
  Mid(tblStr(0), DebChar, 3) = "@@ "
  Mid(tblStr(1), DebChar, 3) = "  @"
  Mid(tblStr(2), DebChar, 3) = "@@ "
  Mid(tblStr(3), DebChar, 3) = "  @"
  Mid(tblStr(4), DebChar, 3) = "@@ "
 Case "4"
  Mid(tblStr(0), DebChar, 3) = "@ @"
  Mid(tblStr(1), DebChar, 3) = "@ @"
  Mid(tblStr(2), DebChar, 3) = "@@@"
  Mid(tblStr(3), DebChar, 3) = "  @"
  Mid(tblStr(4), DebChar, 3) = "  @"
 Case "5"
  Mid(tblStr(0), DebChar, 3) = "@@@"
  Mid(tblStr(1), DebChar, 3) = "@  "
  Mid(tblStr(2), DebChar, 3) = "@@ "
  Mid(tblStr(3), DebChar, 3) = "  @"
  Mid(tblStr(4), DebChar, 3) = "@@ "
 Case "6"
  Mid(tblStr(0), DebChar, 3) = "@@@"
  Mid(tblStr(1), DebChar, 3) = "@  "
  Mid(tblStr(2), DebChar, 3) = "@@@"
  Mid(tblStr(3), DebChar, 3) = "@ @"
  Mid(tblStr(4), DebChar, 3) = "@@@"
 Case "7"
  Mid(tblStr(0), DebChar, 3) = "@@@"
  Mid(tblStr(1), DebChar, 3) = "  @"
  Mid(tblStr(2), DebChar, 3) = " @ "
  Mid(tblStr(3), DebChar, 3) = "@  "
  Mid(tblStr(4), DebChar, 3) = "@  "
 Case "8"
  Mid(tblStr(0), DebChar, 3) = "@@@"
  Mid(tblStr(1), DebChar, 3) = "@ @"
  Mid(tblStr(2), DebChar, 3) = "@@@"
  Mid(tblStr(3), DebChar, 3) = "@ @"
  Mid(tblStr(4), DebChar, 3) = "@@@"
 Case "9"
  Mid(tblStr(0), DebChar, 3) = "@@@"
  Mid(tblStr(1), DebChar, 3) = "@ @"
  Mid(tblStr(2), DebChar, 3) = "@@@"
  Mid(tblStr(3), DebChar, 3) = "  @"
  Mid(tblStr(4), DebChar, 3) = "@@@"
 End Select
End Sub



Private Sub Afficher()
 
 varNbrPtsAfi = 0
 
 For t = 0 To varNbrPtsOrg
  'Récupération des points
  xx = tblPtsOrg(t, 0)
  yy = tblPtsOrg(t, 1)
  zz = tblPtsOrg(t, 2)
  
  varAngRadRotX = DegRad(varAngDegRotX)
  varAngRadRotY = DegRad(varAngDegRotY)
  varAngRadRotZ = DegRad(varAngDegRotZ)
  SinX = Sin(varAngRadRotX)
  CosX = Cos(varAngRadRotX)
  SinY = Sin(varAngRadRotY)
  CosY = Cos(varAngRadRotY)
  SinZ = Sin(varAngRadRotZ)
  CosZ = Cos(varAngRadRotZ)
  'Rotation sur l'axe X  (haut-bas/profondeur)
  yy2 = (yy * CosX) - (zz * SinX)
  zz2 = (yy * SinX) + (zz * CosX)
  yy = yy2: zz = zz2
  'Rotation sur l'axe Y  (gauche-droite/profondeur)
  xx2 = (xx * CosY) - (zz * SinY)
  zz2 = (xx * SinY) + (zz * CosY)
  xx = xx2: zz = zz2
  'Rotation sur l'axe Z  (A plat en face)
  xx2 = (xx * CosZ) - (yy * SinZ)
  yy2 = (xx * SinZ) + (yy * CosZ)
  xx = xx2: yy = yy2
  
  'Décalage dans les plans
  xx = xx + varX
  yy = yy + varY
  zz = zz + varZ
  
  'Si Z<0, disparait (ne l'affiche pas)
  If zz <= 0 Then GoTo FinFor
  
  'Correspondance 2D de la 3D  (Fuyantes)
  ' et mettre le point 0 des fuyantes au centre de l'écran
  xxx = varCX + (xx * (varCoZ ^ zz))
  yyy = varCY - (yy * (varCoZ ^ zz))
  
  'Calcul de la Taille de la "Méta-Ball"
  Txy = varCoTxy * (varCoZ ^ zz)
  'Si la "Méta-Ball" fait 3 fois sa taille, ne l'affiche pas
  If Txy > (50 * 2) Then GoTo FinFor
  'Place la "Méta-Ball" au centre du point (2D)
  xxx = xxx - (Txy \ 2)
  yyy = yyy - (Txy \ 2)
  
  tblPtsAfi(varNbrPtsAfi, 0) = zz
  tblPtsAfi(varNbrPtsAfi, 1) = Txy
  tblPtsAfi(varNbrPtsAfi, 2) = xxx
  tblPtsAfi(varNbrPtsAfi, 3) = yyy
  varNbrPtsAfi = varNbrPtsAfi + 1
  
FinFor:
 Next t
 varNbrPtsAfi = varNbrPtsAfi - 1
 
 
 'Affichage proprement dit
 frmMere.Cls
 For tt = varNbrPtsAfi To 0 Step -1
  z = 0
  ttt = 0
  For t = 0 To varNbrPtsAfi
   If tblPtsAfi(t, 0) > z Then
    z = tblPtsAfi(t, 0)
    ttt = t
   End If
  Next t
  tblPtsAfi(ttt, 0) = 0
  tblPtsOrder(tt) = ttt
 Next tt
 
 For t = 0 To varNbrPtsAfi
  Txy = tblPtsAfi(tblPtsOrder(t), 1)
  xxx = tblPtsAfi(tblPtsOrder(t), 2)
  yyy = tblPtsAfi(tblPtsOrder(t), 3)
  
  'Affichage à proprement parlé
  frmMere.PaintPicture PaintedBalls1.GetBallMaskPicture, xxx, yyy, Txy, Txy, 0, 0, 50, 50, vbSrcAnd
  '16 / 35
  If varModeAffi = 0 Then
   frmMere.PaintPicture PaintedBalls1.GetColoredBall(varTexture, varColorBall), xxx, yyy, Txy, Txy, 0, 0, 50, 50, vbSrcPaint
  Else
   frmMere.PaintPicture PaintedBalls1.GetTexturedBall(varTexture), xxx, yyy, Txy, Txy, 0, 0, 50, 50, vbSrcPaint
  End If
 Next t
End Sub



Private Function RadDeg(ByVal vRad As Single) As Single
 'Pour convertir des radians en degrés, multipliez-les par 180/pi.
 RadDeg = vRad * (180 / PI)
End Function


Private Function DegRad(ByVal vDeg As Single) As Single
 'Pour convertir des degrés en radians, multipliez-les par pi/180.
 DegRad = vDeg * (PI / 180)
End Function
