Attribute VB_Name = "Module1"
Option Explicit
Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Const HKEY_CURRENT_USER = &H80000001

Function GetStartMenuPath$()
 GetStartMenuPath$ = GetRegKey$("Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs")
End Function

Function GetDesktopPath$()
 GetDesktopPath$ = GetRegKey$("Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
End Function

Function GetRegKey$(Path$, Key$)
 Dim Resultat As Long
 Dim Ident As Long
 Dim Donnee As String
 Dim TailleBuffer As Long
 
 Resultat = 0
 Resultat = RegCreateKey(HKEY_CURRENT_USER, Path$, Ident)
 Resultat = RegQueryValueEx(Ident, Key$, 0&, 1, 0&, TailleBuffer)
 If TailleBuffer < 2 Then
  GetRegKey$ = ""
  Exit Function
 End If

 Donnee = String(TailleBuffer + 1, " ")
 Resultat = RegQueryValueEx(Ident, Key$, 0&, 1, ByVal Donnee, TailleBuffer)
 Donnee = Left(Donnee, TailleBuffer - 1)
 GetRegKey$ = Donnee
End Function

