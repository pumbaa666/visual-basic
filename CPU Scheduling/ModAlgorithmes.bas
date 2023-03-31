Attribute VB_Name = "ModAlgorithmes"
Public Type structEntree
    vNom As String
    vArrivee As Integer
    vDuree As Integer
End Type

Public tEntree() As structEntree
Public vNbEntree As Integer

Function Chronogramme()
Dim vCount As Integer

    '************************ Affiche le 1er *************************'
    FrmChrono.LblNom(0).Caption = tEntree(0).vNom
    FrmChrono.LblNom(0).Left = 500
    FrmChrono.LblNom(0).Visible = True

    FrmChrono.ShpPlace(0).Left = 475
    FrmChrono.ShpPlace(0).Visible = True

    FrmChrono.LblArrivee(0).Caption = tEntree(0).vArrivee
    FrmChrono.LblArrivee(0).Left = 500
    FrmChrono.LblArrivee(0).Visible = True
    '*****************************************************************'

    For vCount = 1 To vNbEntree - 1
        Load FrmChrono.LblNom(vCount)
        Load FrmChrono.ShpPlace(vCount)
        Load FrmChrono.LblArrivee(vCount)
        Load FrmChrono.LblFin(vCount)

        FrmChrono.LblFin(vCount).Caption = tEntree(vCount - 1).vArrivee + tEntree(vCount - 1).vDuree
        FrmChrono.LblFin(vCount).Left = 150 + 1000 * vCount
        FrmChrono.LblFin(vCount).Visible = True
    
        If tEntree(vCount).vArrivee < tEntree(vCount - 1).vArrivee + tEntree(vCount - 1).vDuree Then
            tEntree(vCount).vArrivee = tEntree(vCount - 1).vArrivee + tEntree(vCount - 1).vDuree + 1
        End If
    
        FrmChrono.LblNom(vCount).Caption = tEntree(vCount).vNom
        FrmChrono.LblNom(vCount).Left = 500 + 1000 * vCount
        FrmChrono.LblNom(vCount).Visible = True
    
        FrmChrono.ShpPlace(vCount).Left = 475 + 1000 * vCount
        FrmChrono.ShpPlace(vCount).Visible = True
    
        FrmChrono.LblArrivee(vCount).Caption = tEntree(vCount).vArrivee
        FrmChrono.LblArrivee(vCount).Left = 500 + 1000 * vCount
        FrmChrono.LblArrivee(vCount).Visible = True
    Next
    
    Load FrmChrono.LblFin(vCount)
    FrmChrono.LblFin(vCount).Caption = tEntree(vCount - 1).vArrivee + tEntree(vCount - 1).vDuree
    FrmChrono.LblFin(vCount).Left = 150 + 1000 * vCount
    FrmChrono.LblFin(vCount).Visible = True
    
    FrmChrono.Width = 1000 * vCount + 1000
    FrmChrono.Show

    Liste
End Function

Function Decharger()
Dim vCount As Integer

    On Error Resume Next
    For vCount = 1 To vNbEntree - 1
        Unload FrmChrono.LblArrivee(vCount)
        Unload FrmChrono.LblFin(vCount)
        Unload FrmChrono.LblNom(vCount)
        Unload FrmChrono.ShpPlace(vCount)

        Unload FrmListe.Liste(vCount)
        Unload FrmListe.Label(vCount)
    Next
    Unload FrmChrono.LblFin(vNbEntree)
    Unload FrmListe.Liste(vNbEntree)
    Unload FrmListe.Label(vNbEntree)
    FrmListe.Liste(0).Clear
End Function

Function FCFS()
Dim vCount As Integer
Dim vCount2 As Integer
Dim vTemp As structEntree
    
    For vCount2 = 0 To vNbEntree
        For vCount = vCount2 + 1 To vNbEntree - 1 ' 1 pour vCount2 ?
            If tEntree(vCount).vArrivee < tEntree(vCount - 1).vArrivee Then
                vTemp.vArrivee = tEntree(vCount - 1).vArrivee
                vTemp.vDuree = tEntree(vCount - 1).vDuree
                vTemp.vNom = tEntree(vCount - 1).vNom

                tEntree(vCount - 1).vArrivee = tEntree(vCount).vArrivee
                tEntree(vCount - 1).vDuree = tEntree(vCount).vDuree
                tEntree(vCount - 1).vNom = tEntree(vCount).vNom

                tEntree(vCount).vArrivee = vTemp.vArrivee
                tEntree(vCount).vDuree = vTemp.vDuree
                tEntree(vCount).vNom = vTemp.vNom
            End If
        Next
    Next
End Function

Function Liste()
Dim vCount As Integer
Dim vCount2 As Integer
Dim tReste() As Integer

    ReDim tReste(vNbEntree)
    '***************** Création ******************'
    For vCount = 1 To vNbEntree
        Load FrmListe.Liste(vCount)
        Load FrmListe.Label(vCount)

        FrmListe.Liste(vCount).Left = 300 + FrmListe.Liste(vCount).Width * vCount
        FrmListe.Liste(vCount).Visible = True

        FrmListe.Label(vCount).Left = 300 + FrmListe.Liste(vCount).Width * vCount
        FrmListe.Label(vCount).Visible = True
        FrmListe.Label(vCount).Caption = tEntree(vCount - 1).vNom

        tReste(vCount - 1) = tEntree(vCount - 1).vDuree
    Next
    '*********************************************'

    '*********** Remplissage ***********'
    For vCount2 = 0 To tEntree(vNbEntree - 1).vArrivee + tEntree(vNbEntree - 1).vDuree
        FrmListe.Liste(0).AddItem vCount2
        For vCount = 1 To vNbEntree
            If vCount2 >= tEntree(vCount - 1).vArrivee And vCount2 <= tEntree(vCount - 1).vArrivee + tEntree(vCount - 1).vDuree Then
                FrmListe.Liste(vCount).AddItem tReste(vCount - 1)
                tReste(vCount - 1) = tReste(vCount - 1) - 1
            Else
                FrmListe.Liste(vCount).AddItem " - "
            End If
        Next
    Next
    '***********************************'

    FrmListe.Width = 1000 + FrmListe.Liste(1).Width * vCount
    FrmListe.Show
End Function

Function SJFNot()
Dim vCount As Integer
Dim vCount2 As Integer
Dim vTemp As structEntree
Dim vEnCours As Integer
Dim tMin() As Integer
'Dim vNbArrive As Integer
Dim vMin As Integer
Dim vNum As Integer
Dim tStruct2() As structEntree

    FCFS        ' Trie le tableau par ordre d'arrivée

    '***************** Passe tEntree dans vStruct2 *******************'
    ReDim tStruct2(vNbEntree)
    For vCount2 = 0 To vNbEntree - 1
        tStruct2(vCount).vArrivee = tEntree(vCount).vArrivee
        tStruct2(vCount).vDuree = tEntree(vCount).vDuree
        tStruct2(vCount).vNom = tEntree(vCount).vNom
    Next
    '*****************************************************************'

    '****************** Cherche toute les arrivées *******************'
    For vCount2 = 0 To vNbEntree - 1
        vMin = tEntree(vCount2).vArrivee + tEntree(vCount2).vDuree
        vNum = 0
        For vCount = 1 To vNbEntree - 1
            If tStruct2(vCount).vArrivee < tEntree(vCount).vArrivee + tEntree(vCount).vDuree And tStruct2(vCount).vArrivee <> 0 Then    'Cherche la plus petite durée
                If tStruct2(vCount).vDuree < vMin Then
                    vMin = tEntree(vCount).vDuree
                    vNum = vCount
                End If
            End If
        Next

        If vNum <> 0 Then
            tEntree(vCount).vNom = tEntree(vNum).vNom
            tEntree(vCount).vArrivee = tEntree(vNum).vArrivee
            tEntree(vCount).vDuree = tEntree(vNum).vDuree
        End If

        tStruct2(vNum).vArrivee = 0 ' vCount à la place de vNum ?
    Next
    '*****************************************************************'
End Function
