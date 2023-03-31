Attribute VB_Name = "Module1"
Option Explicit
Public appDVD As Excel.Application
Public wbDVD As Excel.Workbook
Public vNbDVDTot As Integer
Public tListe(1 To 5, 0 To 800) As String
Public tNewListe(1 To 5, 0 To 800) As String
Public vTestListe5 As Boolean

Sub Setup()
    On Error Resume Next
    Set appDVD = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ' Si Excel n'est pas ouvert
        Set appDVD = CreateObject("Excel.Application") ' L'ouvre
    End If
    Err.Clear

    On Error GoTo 0
    Set wbDVD = appDVD.Workbooks.Open(App.Path & "\DVD Loïc.xls")
End Sub

Sub CleanUp()
    appDVD.Quit
    Set appDVD = Nothing
    Set wbDVD = Nothing
End Sub

Sub LectureFichier()
Dim vColonne As Integer
Dim vLigne As Integer
Dim shtContinent As Excel.Worksheet
Dim vTemp As String

    Set shtContinent = wbDVD.Sheets("Sheet1")
    vLigne = 1
    vColonne = 1
    While (shtContinent.Cells(vLigne, vColonne) <> "")
'    vNbDVDTot = 9
'    While (vLigne < 10)
        FrmListe.Liste(0).AddItem vLigne
        For vColonne = 1 To 5
           vTemp = shtContinent.Cells(vLigne, vColonne)
           FrmListe.Liste(vColonne).AddItem vTemp
           tListe(vColonne, vLigne) = vTemp
        Next
        vLigne = vLigne + 1
        vColonne = 1
    Wend
    vNbDVDTot = vLigne - 1

    Set shtContinent = Nothing
End Sub

Sub Preter(ByVal vNum As Integer, ByVal vNom As String)
Dim shtPreter As Excel.Worksheet
    Set shtPreter = wbDVD.Sheets("Sheet1")

    shtPreter.Cells(vNum, 5) = vNom
    Set shtPreter = Nothing
End Sub

Sub RefreshListePrete()
Dim vCount As Integer
    vTestListe5 = False
    FrmListe.Liste(5).Clear
    For vCount = 0 To vNbDVDTot
        FrmListe.Liste(5).AddItem tListe(5, vCount)
    Next
End Sub

Sub ToutCacher()
    FrmAdd.Hide
    FrmPreter.Hide
    FrmChercher.Hide
'    If FrmListe.Liste(0).ListCount <> vNbDVDTot Then
'        RefreshListe
'    End If
End Sub

Sub Ajouter(ByVal vTitre As String, ByVal vGenre As String, ByVal vActeurs As String, ByVal vNote As String, ByVal vPrete As String)
Dim shtAjouter As Excel.Worksheet
    Set shtAjouter = wbDVD.Sheets("Sheet1")
    vNbDVDTot = vNbDVDTot + 1
    shtAjouter.Cells(vNbDVDTot, 1) = vTitre
    shtAjouter.Cells(vNbDVDTot, 2) = vGenre
    If vActeurs <> "" Then
        shtAjouter.Cells(vNbDVDTot, 3) = vActeurs
    End If
    If vNote <> "1 --> 9" Then
        shtAjouter.Cells(vNbDVDTot, 4) = vNote
    End If
        shtAjouter.Cells(vNbDVDTot, 5) = vPrete
    Set shtAjouter = Nothing
End Sub

Sub DelFichier(ByVal vNum As Integer)
Dim vCount As Integer
Dim vCntColonne As Integer
Dim shtDel As Excel.Worksheet

    Set shtDel = wbDVD.Sheets("Sheet1")
    For vCount = vNum To vNbDVDTot + 1
        For vCntColonne = 1 To 5
            shtDel.Cells(vCount, vCntColonne) = shtDel.Cells(vCount + 1, vCntColonne)
        Next
    Next
    Set shtDel = Nothing
End Sub

Sub RefreshListe()
Dim vColonne As Integer
Dim vLigne As Integer

    If FrmListe.Liste(0).ListCount <> vNbDVDTot Then
        ClearListe
        For vLigne = 1 To vNbDVDTot
            FrmListe.Liste(0).AddItem vLigne
            For vColonne = 1 To 5
               FrmListe.Liste(vColonne).AddItem tListe(vColonne, vLigne)
               tListe(vColonne, vLigne) = tListe(vColonne, vLigne)
            Next
        Next
    End If
End Sub

Sub ClearListe()
Dim vCount As Integer
    For vCount = 0 To 5
        FrmListe.Liste(vCount).Clear
    Next
End Sub

Sub BouttonPreter()
    ToutCacher
    FrmPreter.Caption = "Mes DVD - Prêter un DVD"
    FrmPreter.Frame1.Visible = True
    FrmPreter.Frame2.Top = 720
    FrmPreter.Frame2.Left = 600
    FrmPreter.Height = 1900
    FrmPreter.Top = 0
    FrmPreter.Left = 5000
    FrmPreter.Show
    FrmPreter.TxtNom.SetFocus
End Sub

Sub BouttonReprendre()
    ToutCacher
    FrmPreter.Caption = "Mes DVD - Reprendre un DVD"
    FrmPreter.Frame1.Visible = False
    FrmPreter.Frame2.Top = 120
    FrmPreter.Height = 1300
    FrmPreter.Frame2.Left = 100
    FrmPreter.Top = 0
    FrmPreter.Left = 5000
    FrmPreter.Show
    FrmPreter.TxtNum.SetFocus
End Sub

Sub DelDVD(ByVal vNum As Integer)
Dim vCount As Integer

    FrmListe.Liste(5).ListIndex = vNum - 1
    For vCount = 0 To 5
        FrmListe.Liste(vCount).RemoveItem vNum - 1
    Next
    vNbDVDTot = vNbDVDTot - 1
    FrmListe.Liste(0).Clear
    For vCount = 1 To vNbDVDTot
        FrmListe.Liste(0).AddItem vCount
    Next
    DelFichier vNum
    FrmListe.Liste(0).ListIndex = vNum - 2
End Sub

Sub BouttonModifier()
    ToutCacher
    FrmAdd.Top = 0
    FrmAdd.Left = 5000
    FrmAdd.Caption = "Mes DVD - Modifier un DVD"
    FrmAdd.Frame1.Top = 720
    FrmAdd.TxtNum.Visible = True
    FrmAdd.LblPrete.Visible = True
    FrmAdd.TxtPrete.Visible = True
    FrmAdd.Height = 3600
    FrmAdd.Show
    FrmAdd.TxtNum.SetFocus
End Sub
