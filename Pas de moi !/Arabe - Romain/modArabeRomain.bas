Attribute VB_Name = "modArabeRomain"
Public Function ArabeRomain(Nombre As Integer) As String

' Test de la validité < --
If Nombre < 1 Or Nombre > 3999 Then
MsgBox "La valeur doit être comprise entre 1 et 3999 inclus.", vbCritical + vbOKOnly, "Erreur"
ArabeRomain = ""
Exit Function
Else
' Remplissage des blancs < --
Dim strNombreArabe As String
strNombreArabe = str(Nombre)
strNombreArabe = Trim(strNombreArabe)
strNombreArabe = String((4 - Len(strNombreArabe)), "0") & strNombreArabe
' -- > Remplissage des blancs
End If
' -- > Test de la validité

' Déclaration < --
Dim strUnitesArabe As String
Dim strDizainesArabe As String
Dim strCentainesArabe As String
Dim strMilliersArabe As String
Dim strUnitesRomain As String
Dim strDizainesRomain As String
Dim strCentainesRomain As String
Dim strMilliersRomain As String
Dim strNombreRomain As String
' -- > Déclarations

' Décomposition du nombre < --
strUnitesArabe = Right(strNombreArabe, 1)
strDizainesArabe = Mid(strNombreArabe, 3, 1)
strCentainesArabe = Mid(strNombreArabe, 2, 1)
strMilliersArabe = Left(strNombreArabe, 1)
' -- > Décomposition du nombre

' Unités < --
If strUnitesArabe = "0" Then
strUnitesRomain = ""
End If
If Not Val(strUnitesArabe) > 3 Or Val(strUnitesArabe) < 1 Then
strUnitesRomain = String(Val(strUnitesArabe), "I")
End If
If strUnitesArabe = 4 Then
strUnitesRomain = "IV"
End If
If Not Val(strUnitesArabe) > 8 Or Val(strUnitesArabe) < 5 Then
On Error Resume Next
strUnitesRomain = "V" & String((Val(strUnitesArabe) - 5), "I")
End If
If Val(strUnitesArabe) = 9 Then
strUnitesRomain = "IX"
End If
' -- > Unités

' Dizaines < --
If strDizainesArabe = "0" Then
strDizainesRomain = ""
End If
If Not Val(strDizainesArabe) > 3 Or Val(strDizainesArabe) < 1 Then
strDizainesRomain = String(Val(strDizainesArabe), "X")
End If
If strDizainesArabe = 4 Then
strDizainesRomain = "XL"
End If
If Not Val(strDizainesArabe) > 8 Or Val(strDizainesArabe) < 5 Then
On Error Resume Next
strDizainesRomain = "L" & String((Val(strDizainesArabe) - 5), "X")
End If
If Val(strDizainesArabe) = 9 Then
strDizainesRomain = "XC"
End If
' -- > Dizaines

' Centaines < --
If strCentainesArabe = "0" Then
strCentainesRomain = ""
End If
If Not Val(strCentainesArabe) > 3 Or Val(strCentainesArabe) < 1 Then
strCentainesRomain = String(Val(strCentainesArabe), "C")
End If
If strCentainesArabe = 4 Then
strCentainesRomain = "CD"
End If
If Not Val(strCentainesArabe) > 8 Or Val(strCentainesArabe) < 5 Then
On Error Resume Next
strCentainesRomain = "D" & String((Val(strCentainesArabe) - 5), "C")
End If
If Val(strCentainesArabe) = 9 Then
strCentainesRomain = "CM"
End If
' -- > Centaines

' Milliers < --
If strMilliersArabe = "0" Then
strMilliersRomain = ""
End If
If Not Val(strMilliersArabe) > 3 Or Val(strMilliersArabe) < 1 Then
strMilliersRomain = String(Val(strMilliersArabe), "M")
End If
If strMilliersArabe > 3 Then
MsgBox "Overflow.", vbCritical + vbOKOnly, "Error"
End If
' -- > Milliers

' Retour < --
strNombreRomain = strMilliersRomain & strCentainesRomain & strDizainesRomain & strUnitesRomain
ArabeRomain = strNombreRomain
' -- > Retour

End Function
