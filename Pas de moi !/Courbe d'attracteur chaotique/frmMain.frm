VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Attracteurs Chaotiques"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Dessine une courbe d'attraction d'après le modèle de Lyapunov
'aussi appelés attracteurs chaotiques ou attracteurs étranges ou
'encore attracteurs quadratiques (2 dimensions, bien qu'on ne s'y
'limite pas ici)

'Fenetre de rendu
Dim lX  As Double    'Bord gauche
Dim lY  As Double    'Bord haut
Dim hX  As Double    'Bord droit
Dim hY  As Double    'Bord bas

Dim xL  As Double    'Bord gauche
Dim yL  As Double    'Bord haut
Dim xH  As Double    'Bord droit
Dim yH  As Double    'Bord bas

Dim X   As Double
Dim Y   As Double

Dim mY  As Double
Dim mX  As Double
Dim rS  As Double
Dim dF  As Double
Dim Xe  As Double
Dim Ye  As Double
Dim T   As Double

Dim xMax    As Double
Dim xMin    As Double
Dim yMax    As Double
Dim yMin    As Double
Dim ordMax  As Long
Dim prvMax  As Long

Dim xNew    As Double
Dim yNew    As Double
Dim xSave   As Double
Dim ySave   As Double

Dim N       As Long     'Itérations
Dim MaxN    As Long     'Itérations Max
Dim L       As Long
Dim NL      As Long     'Nombre de logarithmes
Dim LSum    As Long     'Somme des logarithmes
Dim P       As Long
Dim I       As Double

Dim O       As Long
Dim M       As Long
Dim Dims    As Double
Dim Code    As String   'Code de l'attracteur

Dim A(504)  As Double   'Coefficients
Dim V(99)   As Double   'Nombres aléatoires
Dim XS(499) As Double
Dim Ran     As Double   'Nombre aléatoire

Dim xP      As Double
Dim yP      As Double
Dim dLX     As Double
Dim dLY     As Double
Dim dL2     As Double

Dim cX      As Double
Dim cY      As Double
Dim cW      As Double
Dim cH      As Double
Dim CCCL    As Long

Private Function Shuffle()

    'Prépare et retrouve les nombres aléatoires

    Dim J   As Long

    If V(0) = 0 Then
        For J = 0 To 99
            V(J) = Rnd
        Next J
    End If
        
    J = Int(100 * Ran)
    Ran = V(J)
    V(J) = Rnd
End Function

Private Function GetCoeff()

    'Retrouve les coefficients

    O = 2 + Int((ordMax - 1) * Rnd)
    Code = Chr$(59 + 4 * Dims + O)
    
    M = 1
    
    For I = 1 To Dims
        M = M * (O + I)
    Next I
    
    'Construit le code de l'attracteur
    For I = 1 To M
        Shuffle
        Code = Code & Chr$(65 + Int(25 * Ran))
    Next I
    
    'Retrouve les coefficents du polynome
    For I = 1 To M
        A(I) = (Asc(Mid$(Code$, I + 1, 1)) - 77) / 10
    Next I
End Function

Private Function InitRender()
    
    'Réinitialise la fenetre

    lX = -0.1: lY = -0.1: hX = 1.1: hY = 1.1
    Cls
End Function

Private Function SetParameters()

    'Paramètres initiaux

    'Positions
    X = 0.05
    Y = 0.05
    
    'Décalages
    Xe = X + 0.000001
    Ye = Y
    
    GetCoeff                 'Get coefficients
    
    T = 3
    P = 0
    LSum = 0
    N = 0
    NL = 0
    xMin = 1000000!
    xMax = -xMin
    yMin = xMin
    yMax = xMax

End Function

Private Function Iterate()

    'Recalcule les positions

    xNew = A(1) + X * (A(2) + A(3) * X + A(4) * Y)
    xNew = xNew + Y * (A(5) + A(6) * Y)
    yNew = A(7) + X * (A(8) + A(9) * X + A(10) * Y)
    yNew = yNew + Y * (A(11) + A(12) * Y)
    N = N + 1
End Function

Private Function Display()

    'Cadre l'image et lance le rendu

    If N < 100 Or N > 1000 Then PlotForm
    If X < xMin Then xMin = X
    If X > xMax Then xMax = X
    If Y < yMin Then yMin = Y
    If Y > yMax Then yMax = Y
    PlotForm
End Function

Private Function PlotForm()
    On Error Resume Next

    'Dessine la fractale
    If N = 1000 Then ResizeScreen

    XS(P) = X
    P = (P + 1) Mod 500
    I = (P + 500 - prvMax) Mod 500
    
    If Dims = 1 Then
        '1 dimension
        xP = XS(I)
        yP = xNew
    Else
        'dimensions supérieures
        xP = X
        yP = Y
    End If
    
    'Le point sort du plan
    If (N < 1000) Or (xP <= xL) Or (xP >= xH) Or (yP <= yL) Or (yP >= yH) Then Exit Function
    
    'Centre du plan de dessin
    cX = ScaleWidth \ 2 - (hX - lX) \ 2
    cY = ScaleHeight \ 2 - (hY - lY) \ 2
    
    'Dimensions du dessin
    cW = (hX - lX)
    cH = (hY - lY)
    
    'Intensité
    CCCL = 255 - (N / MaxN * 255)
    
    'Dessin du point
    PSet ((xP - lX) / cW * ScaleWidth, (yP - lY) / cH * ScaleHeight), RGB(CCCL, CCCL, CCCL) 'Plot point on screen
End Function

Private Function ResizeScreen()

    'Recalcule la fenetre de dessin

    Dim msgR    As VbMsgBoxResult

    If Dims = 1 Then
        '1 dimension
        yMin = xMin
        yMax = xMax
    End If
    
    If xMax - xMin < 0.000001 Then
        'On prend plus d'espace
        xMin = xMin - 0.0000005
        xMax = xMax + 0.0000005
    End If
    
    If yMax - yMin < 0.000001 Then
        yMin = yMin - 0.0000005
        yMax = yMax + 0.0000005
    End If
    
    'Dimensions
    mX = 0.1 * (xMax - xMin)
    mY = 0.1 * (yMax - yMin)
    
    'On recalcule la fenetre
    xL = xMin - mX
    xH = xMax + mX
    yL = yMin - mY
    yH = yMax + mY
    
    Refresh
    DoEvents
    
    msgR = MsgBox("Voulez vous relancer le calcul ?", vbYesNo, "Operation terminée")
    
    If msgR = vbNo Then
        T = 4
        Exit Function
    Else
        Cls
    End If

    'Copie de la fenetre
    lX = xL
    lY = yL
    hX = xH
    hY = yH

    DoEvents
End Function

Private Function TestResults()

    'Etudie les résultats

    'La courbe tend vers l'infini
    If Abs(xNew) + Abs(yNew) > 1000000! Then T = 2

    'Calcule les exposants
    CalcLyapunov
    
    If N >= MaxN Then T = 2                                 'Attracteur trouvé !
    If Abs(xNew - X) + Abs(yNew - Y) < 0.000001 Then T = 2  'Attracteur trouvé !
    If N > 100 And L < 0.005 Then T = 2                     'Cycle trouvé !
    
    'Mise à jour des valeurs
    X = xNew
    Y = yNew
End Function

Private Function CalcLyapunov()

    'Calcule l'exposant de Lyapunov

    'Sauvegarde les valeurs
    xSave = xNew
    ySave = yNew
    X = Xe
    Y = Ye
    N = N - 1
    
    'Lance une itération
    Iterate
    
    dLX = xNew - xSave              'DeltaX
    dLY = yNew - ySave              'DeltaY
    dL2 = dLX * dLX + dLY * dLY     'Distance
    
    If CSng(dL2) <= 0 Then Exit Function
    
    'Distances
    dF = 1000000000000# * dL2
    rS = 1 / Sqr(dF)

    'Nouvelles valeurs
    Xe = xSave + rS * (xNew - xSave)
    Ye = ySave + rS * (yNew - ySave)
    
    'Echange des valeurs
    xNew = xSave
    yNew = ySave
    
    If dF > 0 Then
        LSum = LSum + Log(dF)
        NL = NL + 1
    End If
    
    'La fractale est un produit de logarithmes
    L = 0.721347 * LSum / NL
End Function

Private Sub Form_Load()
    Show
    DoEvents
    
    'Lance un calcul
    DrawAttractor 4, 50000, 2, 2.7
End Sub

Private Function DrawAttractor(Prev As Long, nMax As Long, oMax As Long, D As Double)

    'Dessine une courbe d'attraction en utilisant le modèle
    'fractal de Lyapunov
    '
    '   Prev : Nombre d'itérations pour placer un point (>1)
    '   nMax : Nombre d'itérations maximales (>Prev)
    '   oMax : Ordre maximum du polynome (>1)
    '   D : Dimension fractale (1>D>4)

    Cls

    'Copie des paramètres
    ordMax = oMax
    Dims = D
    prvMax = Prev
    MaxN = nMax
    
    Randomize

    'Boucle
ReInit:
    InitRender
ReSetParams:
    SetParameters
ReIter:
    Iterate
    Display
    TestResults
    
    'Pour laisser le programme respirer
    If N Mod 100 = 0 Then DoEvents
    
    Select Case T
        Case 1: GoTo ReInit
        Case 2: GoTo ReSetParams
        Case 3: GoTo ReIter
    End Select

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
