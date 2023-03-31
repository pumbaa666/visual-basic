Attribute VB_Name = "Mod_INI"
'---------------------------------------------------------------------------------------
' Module    : Mod_INI
' DateTime  : 30/12/2004 09:45
' Author    : Gwenael
' Ce module n'est pas de moi, mais comme il est super pratique, je l'ai repris
'---------------------------------------------------------------------------------------

'tout le module n'est pas de moi, mais comme il est très pratique, je l'ai repris

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'
' Pour lire dans un fichier INI
'
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'
'Ecrire dans le fichier .Ini
'
Public Sub EcrireIni(stSection As String, stKey As String, stValeur As String, stFichier As String)
' Lecture d'une valeur dans un fichier INI
' stSection est la partie designée entre crochets ([Option] par exemple)
' stKey est le nom de la clé à récupérer (Last_number=... par exemple)
' stValeur est la valeur à stocker
' stFichier est le fichier à manipuler
WritePrivateProfileString stSection, stKey, stValeur, stFichier
End Sub

'
' Lire le fichier .ini
'
Public Function LireIni(stSection As String, stKey As String, stFichier As String)
' Lecture d'une valeur dans un fichier INI
' stSection est la partie designée entre crochets ([Option] par exemple)
' stKey est le nom de la clé à récupérer (Last_number=... par exemple)
Dim stBuf As String, lgBuf As Long, lgRep As Long
' Mise en place du buffer de lecture
stBuf = Space$(255)
lgBuf = 255
lgRep = GetPrivateProfileString(stSection, stKey, "", stBuf, lgBuf, stFichier)
LireIni = Left$(stBuf, lgRep)
End Function




