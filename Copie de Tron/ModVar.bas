Attribute VB_Name = "ModVar"
Public tCoo(1) As Integer
Public tTerrain(59, 35) As Integer

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

    For vCount = 0 To 59
        For vCount2 = 0 To 35
            tTerrain(vCount, vCount2) = 0
            FrmMain.ShpTerrain(vCount2 * 60 + vCount).FillColor = &H8000000F
         Next
    Next
End Function

Public Function Start()
    If vQui = "Client" Then
        FrmMain.ShpTerrain(0).FillColor = vRemoteCouleur
        FrmMain.ShpTerrain(2159).FillColor = vCouleur
        tCoo(X) = 59
        tCoo(Y) = 35
        vDX = -1
    Else
        FrmMain.ShpTerrain(0).FillColor = vCouleur
        FrmMain.ShpTerrain(2159).FillColor = vRemoteCouleur
        tCoo(X) = 0
        tCoo(Y) = 0
        vDX = 1
    End If
    vDY = 0
    tTerrain(0, 0) = 1
    tTerrain(59, 35) = 1
End Function
