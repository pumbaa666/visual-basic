Attribute VB_Name = "Module1"
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public X(1 To 4) As Long 'coord x
Public Y(1 To 4) As Long 'coord y
Public PV(1 To 4) As Long 'points de vie par robot
Public Pas() 'pas de programme attribu�s
Public NbR As Long 'nombre de robots
Public Nom(1 To 4) As String 'nom des robots
Public NumPas As Long
Public DisRep As Long 'distance de rep�rage
Public NbreBlocs As Long 'nombre d'obstacles � g�n�rer
Public AffichageCteRendu As Boolean 'si on affiche le fichier txt � la fin
Public PtEnMoinsSiPasPossible As Boolean
