Attribute VB_Name = "Mod_ENNEMIS"
Public Function loadENNEMI(chemin As String)
'Chargement des ennemis du jeu
Dim currentENNEMI As Integer
Dim fichierini As String
fichierini = App.Path & chemin & ".prm"
nbrENNEMI = Val(LireIni("ENNEMI_Général", "nbrENNEMI", fichierini))
While currentENNEMI <= nbrENNEMI
currentENNEMI = currentENNEMI + 1
ENNEMI(currentENNEMI).x = Val(LireIni("Ennemi_" & currentENNEMI, "INITX", fichierini))
ENNEMI(currentENNEMI).y = Val(LireIni("Ennemi_" & currentENNEMI, "INITY", fichierini))
ENNEMI(currentENNEMI).type = LireIni("Ennemi_" & currentENNEMI, "type", fichierini)
Wend

End Function

Public Function afficheENNEMI()
  For I = 1 To nbrENNEMI
  If ENNEMI(I).type = "" Then AfficherImage ENNEMIsurf(1), ENNEMIsurfddsd(1), ENNEMI(I).x * 32 - 32 + PosMondeX, (ENNEMI(I).y - 1) * 32 - 32 + PosMondeY, ddRect(0, 0, 0, 0)

  
  Next I
'logique des ennemis

End Function
