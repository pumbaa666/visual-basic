Attribute VB_Name = "ModVar"
Public Const DimX = 59
Public Const DimY = 35

Public tCoo(1) As Integer
Public tTerrain(DimX, DimY) As Integer

Public vDY As Integer
Public vDX As Integer

Public vCouleur As String
Public vRemoteCouleur As String
Public vQui As String

Public Const PortLocal = 1001
Public Const PortDistant = 1002

Public Const X = 0
Public Const Y = 1

Public Function ClearTerrain()
Dim vCount As Integer
Dim vCount2 As Integer

    If FrmMain.OptionGrillage.Checked = True Then
        For vCount = 0 To DimX
            For vCount2 = 0 To DimY
                tTerrain(vCount, vCount2) = 0
                FrmMain.ShpTerrain(vCount2 * (DimX + 1) + vCount).FillColor = &H8000000F
             Next
        Next
    Else
        For vCount = 0 To DimX
            For vCount2 = 0 To DimY
                tTerrain(vCount, vCount2) = 0
                FrmMain.ShpTerrain(vCount2 * (DimX + 1) + vCount).FillColor = &H8000000F
             Next
        Next
    End If
End Function

Public Function Start()
    If vQui = "Client" Then
        FrmMain.ShpTerrain(0).FillColor = vRemoteCouleur
        FrmMain.ShpTerrain((DimX + 1) * (DimY + 1) - 1).FillColor = vCouleur
        tCoo(X) = 59
        tCoo(Y) = 35
        vDX = -1
    Else
        FrmMain.ShpTerrain(0).FillColor = vCouleur
        FrmMain.ShpTerrain((DimX + 1) * (DimY + 1) - 1).FillColor = vRemoteCouleur
        tCoo(X) = 0
        tCoo(Y) = 0
        vDX = 1
    End If
    vDY = 0
    tTerrain(0, 0) = 1
    tTerrain(DimX, DimY) = 1
End Function
