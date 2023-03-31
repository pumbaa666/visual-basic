Attribute VB_Name = "Mod_INI"
Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'
' Pour lire dans un fichier INI
'
Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString _
As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'
'Ecrire dans le fichier .Ini
'
Public Sub EcrireIni(stSection As String, stKey As String, stValeur As String, stFichier As String)
' stSection est la partie design�e entre crochets ([Option] par exemple)
' stKey est le nom de la cl� � r�cup�rer (Last_number=... par exemple)
' stValeur est la valeur � stocker
' stFichier est le fichier � manipuler
WritePrivateProfileString stSection, stKey, stValeur, stFichier
End Sub

'
' Lire le fichier .ini
'
Public Function LireIni(stSection As String, stKey As String, stFichier As String) As String
' stSection est la partie design�e entre crochets ([Option] par exemple)
' stKey est le nom de la cl� � r�cup�rer (Last_number=... par exemple)
Dim stBuf As String, lgBuf As Long, lgRep As Long
' Mise en place du buffer de lecture
stBuf = Space$(255)
lgBuf = 255
lgRep = GetPrivateProfileString(stSection, stKey, "", stBuf, lgBuf, stFichier)
LireIni = Left$(stBuf, lgRep)
End Function




