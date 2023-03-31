Attribute VB_Name = "Module"

'############################
'#MODULE DE GESTION GENERALE#
'############################

Public TamponTexte As String
Public NumVersion As String

'Variables pour le redimensionnement
Public Longueur As Integer

'APIs et constantes pour le resize et déplacement
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()

Public Sub ReMap()

With Main
    'On redimmensionne correctement les contrôles
    .LabelTitle.Width = .Width
    .ShapeTitle.Width = .Width
    .ButtonQuitter.Left = .ScaleWidth - .ButtonQuitter.Width - Longueur + 10
    .ButtonAgrandir.Left = .ButtonQuitter.Left - 240
    .ButtonRéduire.Left = .ButtonAgrandir.Left - 240
    .CoolBar.Width = .Width
    .Text.Move 0, .CoolBar.Height + .ShapeTitle.Height - 30, .Width, .Height - .Label.Height - .CoolBar.Height - .ShapeTitle.Height + 50
    .Label.Move 0, .Text.Height + .CoolBar.Height + .ShapeTitle.Height - 50, .Width, .Label.Height
    .Progress.Top = .Label.Top
    .Progress.Left = .Width - .Progress.Width - .ImageResize.Width
    .ImageResize.Move .Width - .ImageResize.Width - 10, .Label.Top
    .ShapeResize.Move .ImageResize.Left, .ImageResize.Top
    'Sauvegarde des dimensions
    SaveSetting "Pyro-Notes III", "Config", "Width", .Width
    SaveSetting "Pyro-Notes III", "Config", "Height", .Height
End With

End Sub

Public Sub LoadParams()

'Chargement des paramètres

With Main
    If GetSetting("Pyro-Notes III", "AssignTXT", "Notepad") = "" Then
        Config.OptionNotepadTXT.Value = True
    Else
        Config.OptionPN3TXT.Value = True
    End If
    If GetSetting("Pyro-Notes III", "AssignRTF", "Wordpad") = "" Then
        Config.OptionWordpadRTF.Value = True
    Else
        Config.OptionPN3RTF.Value = True
    End If
    .Move GetSetting("Pyro-Notes III", "Config", "Left", .Left), GetSetting("Pyro-Notes III", "Config", "Top", .Top), GetSetting("Pyro-Notes III", "Config", "Width", .Width), GetSetting("Pyro-Notes III", "Config", "Height", .Height)
    .Text.Font.Name = GetSetting("Pyro-Notes III", "Config", "Font")
    .Text.BackColor = GetSetting("Pyro-Notes III", "Config", "BackColor")
    If GetSetting("Pyro-Notes III", "Config", "Bold") = "1" Then .Text.Font.Bold = True
    If GetSetting("Pyro-Notes III", "Config", "Bold") = "0" Then .Text.Font.Bold = False
    If GetSetting("Pyro-Notes III", "Config", "Italic") = "1" Then .Text.Font.Italic = True
    If GetSetting("Pyro-Notes III", "Config", "Italic") = "0" Then .Text.Font.Italic = False
    If GetSetting("Pyro-Notes III", "Config", "Underline") = "1" Then .Text.Font.Underline = True
    If GetSetting("Pyro-Notes III", "Config", "Underline") = "0" Then .Text.Font.Underline = False
    If GetSetting("Pyro-Notes III", "Config", "Strikethrough") = "1" Then .Text.Font.Strikethrough = True
    If GetSetting("Pyro-Notes III", "Config", "Strikethrough") = "0" Then .Text.Font.Strikethrough = False
    If GetSetting("Pyro-Notes III", "Config", "Size") = "1" Then .Text.Font.Size = True
    If GetSetting("Pyro-Notes III", "Config", "Size") = "0" Then .Text.Font.Size = False
End With

End Sub

Public Sub SaveBaseParams()

'Sauvegarde des paramètres de départ

With Main
    SaveSetting "Pyro-Notes III", "Config", "Left", .Left
    SaveSetting "Pyro-Notes III", "Config", "Top", .Top
    SaveSetting "Pyro-Notes III", "Config", "Width", .Width
    SaveSetting "Pyro-Notes III", "Config", "Height", .Height
    SaveSetting "Pyro-Notes III", "Config", "Font", .Text.Font.Name
    SaveSetting "Pyro-Notes III", "Config", "BackColor", .Text.BackColor
    If .Text.Font.Bold = True Then SaveSetting "Pyro-Notes III", "Config", "Bold", "1"
    If .Text.Font.Bold = False Then SaveSetting "Pyro-Notes III", "Config", "Bold", "0"
    If .Text.Font.Italic = True Then SaveSetting "Pyro-Notes III", "Config", "Italic", "1"
    If .Text.Font.Italic = False Then SaveSetting "Pyro-Notes III", "Config", "Italic", "0"
    If .Text.Font.Underline = True Then SaveSetting "Pyro-Notes III", "Config", "Underline", "1"
    If .Text.Font.Underline = False Then SaveSetting "Pyro-Notes III", "Config", "Underline", "0"
    If .Text.Font.Strikethrough = True Then SaveSetting "Pyro-Notes III", "Config", "Strikethrough", "1"
    If .Text.Font.Strikethrough = False Then SaveSetting "Pyro-Notes III", "Config", "Strikethrough", "0"
    If .Text.Font.Size = True Then SaveSetting "Pyro-Notes III", "Config", "Size", "1"
    If .Text.Font.Size = False Then SaveSetting "Pyro-Notes III", "Config", "Size", "0"
End With

End Sub

Public Sub Status(Text As String)

'Changement du message dans la barre de status

With Main
    .Label.Caption = " " & Text
End With

End Sub

Public Sub LoadPolices()

'Dim Incr As Integer

'Chargement de la liste des polices

'ReDim TabFonts(Screen.FontCount - 1)
'For Incr = 0 To Screen.FontCount - 1
'    TabFonts(Incr) = Screen.Fonts(Incr)
'Next

End Sub

Public Sub LockMain()

'Fonction de verrouillage de Main lors d'un chargement

With Main
    .CoolBar.Enabled = False
    .Text.BackColor = &HE0E0E0
    .Text.Enabled = False
End With

End Sub

Public Sub UnlockMain()

'Fonction de déverrouillage de Main lors d'un chargement

With Main
    .CoolBar.Enabled = True
    .Text.BackColor = &H80000018
    .Text.Enabled = True
End With

End Sub
