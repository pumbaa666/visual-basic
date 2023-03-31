Attribute VB_Name = "Module1"
Public vNom As String
Public vDate As Date
Public vMailAnni As String
    
Function fSupprimer(ByVal vPlace As Integer)
Dim vData As String
Dim vCount As Integer
Dim vCount2 As Integer

    If MsgBox("Voulez-vous vraiment supprimer cette persone?", vbYesNo, "Suppression") = vbYes Then
        Open "c:\temp\donnees.dat" For Input As #1
        Open "c:\temp\donnees.tmp" For Output As #2
        Do
            Line Input #1, vData
            If vCount2 <> vPlace Then
                Print #2, vData
            End If
            vCount = vCount + 1
            If vCount = 4 Then
                vCount2 = vCount2 + 1
                vCount = 0
            End If
        Loop Until (EOF(1))
        Close #1
        Close #2
        
        FileCopy "c:\temp\donnees.tmp", "c:\temp\donnees.dat"
        Kill "c:\temp\donnees.tmp"
        MsgBox "Suppression correctement effectuée", vbInformation
    End If
End Function
