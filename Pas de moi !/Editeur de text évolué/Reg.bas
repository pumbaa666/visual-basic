Attribute VB_Name = "Reg"

'##########################################
'#MODULE DE GESTION DE LA BASE DE REGISTRE#
'##########################################

'Déclarations d'APIs et de constantes pour la base de registre
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKEY As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BIN = 3
Const REG_DWORD = 4

Public Sub AssignNotepadTXT()

'Assignation des fichiers TXT à Notepad

If GetSetting("Pyro-Notes III", "AssignTXT", "Notepad") <> "" Then
    SaveRegString HKEY_CLASSES_ROOT, "txtfile\shell\open\command", GetSetting("Pyro-Notes III", "AssignTXT", "Notepad"), REG_EXPAND_SZ
    SaveRegString HKEY_CLASSES_ROOT, "txtfile\DefaultIcon", GetSetting("Pyro-Notes III", "AssignTXT", "DefaultIcon"), REG_EXPAND_SZ
    SaveRegString HKEY_CLASSES_ROOT, "txtfile", GetSetting("Pyro-Notes III", "AssignTXT", "Description"), REG_EXPAND_SZ
    DeleteSetting "Pyro-Notes III", "AssignTXT", "Notepad"
    DeleteSetting "Pyro-Notes III", "AssignTXT", "DefaultIcon"
    DeleteSetting "Pyro-Notes III", "AssignTXT", "Description"
End If

End Sub

Public Sub AssignPN3TXT()

Dim Chemin As String

'Enregistrement des clés de base
If GetSetting("Pyro-Notes III", "AssignTXT", "Notepad") = "" Then
    SaveSetting "Pyro-Notes III", "AssignTXT", "Notepad", GetRegString(HKEY_CLASSES_ROOT, "txtfile\shell\open\command", "", REG_EXPAND_SZ)
    SaveSetting "Pyro-Notes III", "AssignTXT", "DefaultIcon", GetRegString(HKEY_CLASSES_ROOT, "txtfile\DefaultIcon", "", REG_EXPAND_SZ)
    SaveSetting "Pyro-Notes III", "AssignTXT", "Description", GetRegString(HKEY_CLASSES_ROOT, "txtfile", "", REG_EXPAND_SZ)
End If

'Assignation des fichiers TXT à PN3

If Right(App.Path, 1) = "\" Then
    Chemin = App.Path
Else
    Chemin = App.Path & "\"
End If
SaveRegString HKEY_CLASSES_ROOT, "txtfile\shell\open\command", Chemin & App.EXEName & ".exe" & " %1", REG_SZ
SaveRegString HKEY_CLASSES_ROOT, "txtfile\DefaultIcon", Chemin & App.EXEName & ".exe,1", REG_SZ
SaveRegString HKEY_CLASSES_ROOT, "txtfile", "Fichier Pyro-Notes", REG_SZ

End Sub

Public Sub AssignWordpadRTF()

'Assignation des fichiers RTF à Wordpad

If GetSetting("Pyro-Notes III", "AssignRTF", "Wordpad") <> "" Then
    SaveRegString HKEY_CLASSES_ROOT, "rtffile\shell\open\command", GetSetting("Pyro-Notes III", "AssignRTF", "Wordpad"), REG_SZ
    SaveRegString HKEY_CLASSES_ROOT, "rtffile\DefaultIcon", GetSetting("Pyro-Notes III", "AssignRTF", "DefaultIcon"), REG_SZ
    SaveRegString HKEY_CLASSES_ROOT, "rtffile", GetSetting("Pyro-Notes III", "AssignRTF", "Description"), REG_SZ
    DeleteSetting "Pyro-Notes III", "AssignRTF", "Wordpad"
    DeleteSetting "Pyro-Notes III", "AssignRTF", "DefaultIcon"
    DeleteSetting "Pyro-Notes III", "AssignRTF", "Description"
End If

End Sub

Public Sub AssignPN3RTF()

Dim Chemin As String

'Enregistrement des clés de base
If GetSetting("Pyro-Notes III", "AssignRTF", "Wordpad") = "" Then
    SaveSetting "Pyro-Notes III", "AssignRTF", "Wordpad", GetRegString(HKEY_CLASSES_ROOT, "rtffile\shell\open\command", "", REG_SZ)
    SaveSetting "Pyro-Notes III", "AssignRTF", "DefaultIcon", GetRegString(HKEY_CLASSES_ROOT, "rtffile\DefaultIcon", "", REG_SZ)
    SaveSetting "Pyro-Notes III", "AssignRTF", "Description", GetRegString(HKEY_CLASSES_ROOT, "rtffile", "", REG_SZ)
End If

'Assignation des fichiers RTF à PN3

If Right(App.Path, 1) = "\" Then
    Chemin = App.Path
Else
    Chemin = App.Path & "\"
End If
SaveRegString HKEY_CLASSES_ROOT, "rtffile\shell\open\command", Chemin & App.EXEName & ".exe" & " %1", REG_SZ
SaveRegString HKEY_CLASSES_ROOT, "rtffile\DefaultIcon", Chemin & App.EXEName & ".exe,3", REG_SZ
SaveRegString HKEY_CLASSES_ROOT, "rtffile", "Fichier Pyro-Notes", REG_SZ

End Sub

Private Function GetRegString(HKEY As Long, SubKey As String, Key As String, RegType As Long) As String

Dim KeyHand As Long
Dim BufSize As Long
Dim Buffer As String

'Lecture d'une chaîne dans la base de registre

RegOpenKey HKEY, SubKey, KeyHand
RegQueryValueEx KeyHand, Key, 0, RegType, ByVal 0&, BufSize
Buffer = String(BufSize, " ")
RegQueryValueEx KeyHand, Key, 0, RegType, ByVal Buffer, BufSize
If InStr(Buffer, Chr$(0)) > 0 Then
    GetRegString = Left(Buffer, InStr(Buffer, Chr$(0)))
Else
    GetRegString = Buffer
End If
RegCloseKey KeyHand

End Function

Private Sub SaveRegString(HKEY As String, SubKey As String, Value As String, RegType As Long)

Dim KeyHand As Long

'Enregistrement d'une chaîne dans la base de registre

RegCreateKey HKEY, SubKey, KeyHand
RegSetValueEx KeyHand, "", 0, RegType, ByVal Value, Len(Value)
RegCloseKey KeyHand

End Sub

