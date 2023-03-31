Attribute VB_Name = "MdlStructure"
Public Type Structure
    vTitre As String * 80
    vHeure As String * 8
    vJours(6) As Integer
End Type

Public vNbEnreg As String
