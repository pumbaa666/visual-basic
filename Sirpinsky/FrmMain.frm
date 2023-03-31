VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Sirpinsky"
   ClientHeight    =   11805
   ClientLeft      =   480
   ClientTop       =   390
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   11805
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Carre 
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tPascal(100, 100) As Double

Private Sub Form_Load()
Dim vNbCarre As Integer

    vNbCarre = CreerGrille
    Pascal
    Remplissage (vNbCarre)
End Sub

Private Function CreerGrille() As Integer
Dim i As Double
Dim vNbCarre As Integer

    i = 1
    Do
        Load Carre(i)
        Carre(i).Visible = True
        Carre(i).Left = Carre(i - 1).Left + Carre(i - 1).Width
        Carre(i).Top = Carre(i - 1).Top
        If Carre(i).Left + 2 * Carre(i).Width >= FrmMain.Width Then
            If vCentre = 0 Then
                vCentre = i / 2
            End If
            Carre(i).Left = Carre(0).Left
            Carre(i).Top = Carre(i - 1).Top + Carre(i - 1).Height
            If Carre(i).Top + 2 * Carre(i).Height >= FrmMain.Height Then
                Unload Carre(i)
                vNbCarre = i
                i = -1
            End If
        End If
        i = i + 1
    Loop While (i <> 0)
    CreerGrille = vNbCarre
End Function

Private Function Pascal()
Dim i As Integer
Dim j As Integer

    tPascal(0, 0) = 1
    tPascal(1, 0) = 1
    tPascal(1, 1) = 1

    For i = 1 To 100
        tPascal(i, 0) = 1
        For j = 1 To i - 1
            tPascal(i, j) = tPascal(i - 1, j - 1) + tPascal(i - 1, j)
            tPascal(i, j) = tPascal(i, j) Mod 2 + 2
        Next
        tPascal(i, j) = 1
    Next
    i = 0
End Function

Private Function Remplissage(ByVal vNbCarre As Integer)
Dim i As Integer
Dim j As Integer

    For i = 0 To 100
        For j = 0 To 100
            If 100 * i + j < vNbCarre Then
                If tPascal(i, j) = 2 Or tPascal(i, j) = 1 Then
                    Carre((100 * i + j)).FillColor = &H80000012
                End If
            End If
        Next
    Next
End Function
