Attribute VB_Name = "tools"
'*************************************************************
'* source finalisé en décembre 2004 par madbob
'*************************************************************

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Public Type Matrice_RVB
    R As Integer
    V As Integer
    B As Integer
    x As Integer
End Type

Public Type M_RVB
    R As Byte
    V As Byte
    B As Byte
    x As Byte
    Init As Integer
End Type

'* Api Externe
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

'* variable globales
Public PicBitsT() As Byte, PicInfoT As BITMAP  'objet mmémoire pour le timer à coder
Public PicBits() As Byte, PicInfo As BITMAP
Public PicBits2() As Byte, PicInfo2 As BITMAP
Public Cnt As Long, BytesPerLine As Long
Public Mat As Matrice_RVB
Public Mat_Sav As M_RVB
Public G_TABR() As Byte, G_TABV() As Byte, G_TABB() As Byte
Public G_NTABR() As Byte, G_NTABV() As Byte, G_NTABB() As Byte
Public G_Large As Long, G_Haut As Long, G_Limite As Long
Public G_Pause As Boolean
Public M1_Sav As M_RVB
Public M2_Sav As M_RVB
Public M3_Sav As M_RVB
Public G_Permut As Integer
Public G_Time As Integer
Public G_Pas As Integer
Public G_hdl As Integer

'***********************************************************************************************************************************************

Sub Init(P_Form As Form)
    G_Pause = False
    
    '* Récup par l'application (ici par l'application propriété picture!!!)
    GetObject P_Form.Picture1.Image, Len(PicInfo), PicInfo
    G_Large = PicInfo.bmWidth
    G_Haut = PicInfo.bmHeight
    
    '* Définition dinamique de la taille des tableaux
    ReDim PicBits(1 To (PicInfo.bmWidth * PicInfo.bmHeight * 4)) As Byte
    ReDim PicBitsT(1 To (PicInfo.bmWidth * PicInfo.bmHeight * 4)) As Byte
    ReDim PicBits2(1 To (PicInfo.bmWidth * PicInfo.bmHeight * 4)) As Byte
    
    Call ChargeSav(P_Form)
    
    '* Copy the bitmapbits to the array
    GetBitmapBits P_Form.Picture1.Image, UBound(PicBits), PicBits(1)
    
    '* Préparation des espaces de travail en mémoire
    Call REDIM_GTAB_BVR
End Sub

Sub ChargeSav(P_Form As Form)
    GetBitmapBits P_Form.Picture1.Image, UBound(PicBitsT), PicBitsT(1)
End Sub

'* retourne le centre d'une bmp sur sur une une des composante RVB !!!
'* Pour être utilisable les composantes des pixels doivent être dissociées
'* un tableau pour R, un tableau pour B et un tableau pour V
'* L'axe référence le centre commun à ces 3 tableaux
'* l'axe est calculé à partir de la diagonal il est de type pair ou impair
Function CalculAxeBmp(P_Larg As Long, P_Haut As Long, P_Axe As Long)
Dim L_Pair As Integer

    L_Pair = 0
    P_Axe = (((P_Larg * 2) - 1) - P_Larg) * (P_Larg / 2)
    If P_Larg = P_Haut Then
        If (P_Axe * 2 + 1) = (P_Larg * P_Haut) Then
            L_Pair = 1
        End If
    Else
        If (P_Axe * 2 + 1) = (P_Larg * P_Haut) Then
            L_Pair = 2
        Else
            L_Pair = 2
        End If
    End If

    CalculAxeBmp = L_Pair
End Function

Function MonModBonifacio(ByVal Val1 As Integer, ByVal Val2 As Integer)
Dim Tot As Integer
Dim sTot As Long
    sTot = ((Val1 + Val2) Mod 255)
    Tot = sTot
    MonModBonifacio = Tot
End Function
Function Mondecalage(ByVal Val1 As Integer, ByVal P_Decal As Integer)
If Val1 - P_Decal < 0 Then
    Mondecalage = 0
ElseIf Val1 - P_Decal > 255 Then
    Mondecalage = 255
Else
    Mondecalage = Val1 - P_Decal
End If
End Function

Function Monpythagore(ByVal Val1 As Integer, ByVal Val2 As Integer)
Dim Tot As Integer
Dim sTot As Double
Dim l_v1 As Double
Dim l_v2 As Double

l_v1 = Val1
l_v2 = Val2
    
    sTot = (l_v1 * l_v1) + (l_v2 * l_v2)
    
    Tot = (Sqr(sTot)) Mod 255
    Monpythagore = Tot
End Function

Sub TrtRotationMem(P_Pos As Long, P_TypeDecalH As Long, P_TypeDecalV As Long, _
P_NbDecal As Long, MaxLarg As Long, MaxHaut As Long)
Dim cpt As Long
Dim Limite As Long

Limite = MaxLarg * MaxHaut
cpt = 0

'* boucle Horizontale sens +
For cpt = 1 To P_NbDecal
    If P_Pos + P_TypeDecalH > Limite Then
        P_Pos = P_Pos + P_TypeDecalH Mod P_TypeDecalH
        Exit For
    End If
    
    If cpt = 1 Then
        '* Sauvegarde de la suivante
        M1_Sav.R = G_TABR(P_Pos + P_TypeDecalH)
        M1_Sav.V = G_TABV(P_Pos + P_TypeDecalH)
        M1_Sav.B = G_TABB(P_Pos + P_TypeDecalH)
    
        G_TABR(P_Pos + P_TypeDecalH) = G_TABR(P_Pos)
        G_TABV(P_Pos + P_TypeDecalH) = G_TABV(P_Pos)
        G_TABB(P_Pos + P_TypeDecalH) = G_TABB(P_Pos)
    
    Else
        M2_Sav.R = G_TABR(P_Pos + P_TypeDecalH)
        M2_Sav.V = G_TABV(P_Pos + P_TypeDecalH)
        M2_Sav.B = G_TABB(P_Pos + P_TypeDecalH)
                
        G_TABR(P_Pos + P_TypeDecalH) = M1_Sav.R
        G_TABV(P_Pos + P_TypeDecalH) = M1_Sav.V
        G_TABB(P_Pos + P_TypeDecalH) = M1_Sav.B
        
        M1_Sav.R = M2_Sav.R
        M1_Sav.V = M2_Sav.V
        M1_Sav.B = M2_Sav.B
        
    End If
            
    '* gestion du déclage des positions Horizontales et négatives
    P_Pos = P_Pos + P_TypeDecalH
    
Next cpt

'* boucle Verticale sens + attention à l'incrément en sortant du for précédent
For cpt = 1 To P_NbDecal
    
    If P_Pos + P_TypeDecalV > Limite Then
        Exit For
    End If
    
    M2_Sav.R = G_TABR(P_Pos + P_TypeDecalV)
    M2_Sav.V = G_TABV(P_Pos + P_TypeDecalV)
    M2_Sav.B = G_TABB(P_Pos + P_TypeDecalV)
        
    '* affectation
    G_TABR(P_Pos + P_TypeDecalV) = M1_Sav.R
    G_TABV(P_Pos + P_TypeDecalV) = M1_Sav.V
    G_TABB(P_Pos + P_TypeDecalV) = M1_Sav.B
        
    M1_Sav.R = M2_Sav.R
    M1_Sav.V = M2_Sav.V
    M1_Sav.B = M2_Sav.B
    
    '* gestion du déclage des positions Verticales et positives
    P_Pos = P_Pos + P_TypeDecalV
Next cpt

'*******************
'* boucle Horizontale sens - !!! attention à l'incrément en sortant du for précédent
'P_Pos = P_Pos - P_TypeDecalV
For cpt = 1 To P_NbDecal
    
    If P_Pos - P_TypeDecalH < 1 Then
        'P_Pos = -1 * ((P_Pos - P_TypeDecalH) Mod P_TypeDecalH)
        Exit For
    End If
    
    M2_Sav.R = G_TABR(P_Pos - P_TypeDecalH)
    M2_Sav.V = G_TABV(P_Pos - P_TypeDecalH)
    M2_Sav.B = G_TABB(P_Pos - P_TypeDecalH)
        
    '* affectation
    G_TABR(P_Pos - P_TypeDecalH) = M1_Sav.R
    G_TABV(P_Pos - P_TypeDecalH) = M1_Sav.V
    G_TABB(P_Pos - P_TypeDecalH) = M1_Sav.B
        
    M1_Sav.R = M2_Sav.R
    M1_Sav.V = M2_Sav.V
    M1_Sav.B = M2_Sav.B
        
    '* gestion du décalage des positions Verticales et positives
    P_Pos = P_Pos - P_TypeDecalH
    
Next cpt

'* boucle vertivale sens - !!! attention à l'incrément en sortant du for précédent
For cpt = 1 To P_NbDecal
    If (P_Pos - P_TypeDecalV) < 1 Then
        Exit For
    End If
    
    '* sauvegarde de la valeur cible
    M2_Sav.R = G_TABR(P_Pos - P_TypeDecalV)
    M2_Sav.V = G_TABV(P_Pos - P_TypeDecalV)
    M2_Sav.B = G_TABB(P_Pos - P_TypeDecalV)
        
    '* affectation de la valeur source
    G_TABR(P_Pos - P_TypeDecalV) = M1_Sav.R
    G_TABV(P_Pos - P_TypeDecalV) = M1_Sav.V
    G_TABB(P_Pos - P_TypeDecalV) = M1_Sav.B
        
    '* transfert sauvegarde dans prochaine source
    M1_Sav.R = M2_Sav.R
    M1_Sav.V = M2_Sav.V
    M1_Sav.B = M2_Sav.B
    
    '* gestion du décalage des positions Verticales et positives
    P_Pos = P_Pos - P_TypeDecalV
Next cpt

End Sub

Sub Rotation1PixSurAxe(P_Larg As Long, P_Haut As Long)
Dim L_PasDiag As Long
Dim L_Pair As Integer
Dim L_Axe As Long
Dim L_DecalHoriz As Long
Dim L_DecalVert As Long
Dim P_Sens As Integer
Dim CptAxe As Long
Dim CptPermut As Long
Dim NbPermut As Long
Dim NbFixePermut As Long
Dim L_Sauv As Byte
Dim Position As Long

L_PasDiag = 0
L_Pair = 0
L_Axe = 0
P_Sens = 0
Position = 0
L_DecalHoriz = 1 '* 2
L_DecalVert = P_Larg '* 2

'* initialisations
Call CalculPasDiag(P_Larg, L_PasDiag)
L_Pair = CalculAxeBmp(P_Larg, P_Haut, L_Axe)

'* surcharge cas des axes pair ou impair (point de départ des décalages)
NbFixePermut = 1


    '* Boucle principale -> on parcourt la diagonale à partir de son axe
    '* en revenant au point de départ => on effectue un carré !
    '* On sauve la premire valeur
    M3_Sav.R = G_TABR(L_Axe)
    M3_Sav.V = G_TABV(L_Axe)
    M3_Sav.B = G_TABB(L_Axe)
    For CptAxe = L_Axe To 1 Step -(L_PasDiag)
        Position = CptAxe
        '* debut d'un cycle de permutation
        Mat_Sav.Init = 1
       
        '* Permutation Horizontale sens +
        Call TrtRotationMem(Position, L_DecalHoriz, L_DecalVert, NbFixePermut, P_Larg, P_Haut)
       
        '* Mise à jour des compteurs de permutation
        NbFixePermut = NbFixePermut + 2
    Next CptAxe
    
    '* on insere la dernière valeur sauvée au départ de la spirale
    G_TABR(1) = M3_Sav.R
    G_TABV(1) = M3_Sav.V
    G_TABB(1) = M3_Sav.B
End Sub

Sub REDIM_GTAB_BVR()
    
    '* Redimensionnement dynamique de la mémoire de travail à utiliser
    ReDim G_TABR(1 To (UBound(PicBits)) / 4) As Byte
    ReDim G_TABV(1 To (UBound(PicBits)) / 4) As Byte
    ReDim G_TABB(1 To (UBound(PicBits)) / 4) As Byte

End Sub

'* initialisation du déplacement sur l'axe de la première diagonale
Sub CalculPasDiag(P_Larg As Long, P_PasDiag As Long)
    P_PasDiag = (P_Larg + 2) - 1
End Sub

Sub Dematrisation_GTAB_BVR()
Dim k As Long
Dim cpt As Long
k = 1
Cnt = 1
'* chargement en mémoire dans les tableaux de travail
'* on fait l'économie d'un tableau car image en 24 bits !!! et non 32
For cpt = 1 To (UBound(PicBits)) Step 4
    
    If k > UBound(G_TABB) Then
        Exit For
    End If
    G_TABB(k) = PicBits(cpt)
    G_TABV(k) = PicBits(cpt + 1)
    G_TABR(k) = PicBits(cpt + 2)
    
    k = k + 1
Next cpt
End Sub

Private Function Decaleligneplusun(ByVal L_Larg As Long, ByVal L_Haut As Long)
Dim Decal As Long
Dim T_Ligne As Long
Dim L_Compt As Long
Dim k As Long

Decal = 0
T_Ligne = 0
k = 0
L_Compt = 0

'* Affectation dynamique de la mémoire
ReDim G_TABR(1 To L_lastbytvalide / 4) As Byte
ReDim G_TABV(1 To L_lastbytvalide / 4) As Byte
ReDim G_TABB(1 To L_lastbytvalide / 4) As Byte
        
    '* chargement en mémoire dans les tableaux de travail
    For Cnt = 1 To L_lastbytvalide Step 4
        G_TABB(k) = PicBits(Cnt)
        G_TABV(k) = PicBits(Cnt + 1)
        G_TABR(k) = PicBits(Cnt + 2)
            
        k = k + 1
    Next Cnt
        
    '* Decalage 1 ligne = NBre de pixel en hauteur
    ' <=> Nbre d'indice du tableau / largeur de l'image
    Decal = 1
    L_Compt = k - 1
    k = 1
                
    L_Deb = 1
    L_debTmp = L_Deb
    L_Fin = L_Larg
    L = 1
        
    '* traitement ligne par ligne
    For T_Ligne = 1 To L_Haut
            
        '* on décale tous les élements de la ligne
        '* => re déterminer les indices de travail
            
        For L_Deb = L To (L_Fin)
            '* permutation directe
            '* En cas de rupture de ligne permutation de l'extémité
            v_Decal = L_Deb + (Decal Mod L_Larg)
                 
            '* les indices de fin passent au début
            If v_Decal > L_Fin Then
                v_Decal = L_debTmp
                L_debTmp = L_debTmp + 1
            End If
            PicBits(k) = G_TABB(v_Decal)
            PicBits(k + 1) = G_TABR(v_Decal)
            PicBits(k + 2) = G_TABV(v_Decal)
                
            '* gestion des indices du tableau en 32 bits
            k = k + 4
        Next L_Deb
        '* gestion de la ligne
        L = L_Deb
        L_debTmp = L_Deb - 1
        L_Fin = L_Fin + L_Larg
            
        '* gestion du décalage
        If Decal = (L_Larg - 1) Then
            Decal = 1
        Else
            Decal = (Decal) + 1
        End If
    Next T_Ligne
        
    '* mise à jour graphique
    'SetBitmapBits Me.Picture, UBound(PicBits), PicBits(1)
    'Me.Refresh
    
    '* sortie
    Decaleligneplusun = 0
End Function

Sub RESTIT_MEM()
'* restitution
Dim Cpt_M As Long
Dim Cpt_P As Long
k = 1
Cpt_P = 1

For Cpt_M = 1 To UBound(G_TABB)
    PicBits(Cpt_P) = G_TABB(Cpt_M)
    PicBits(Cpt_P + 1) = G_TABV(Cpt_M)
    PicBits(Cpt_P + 2) = G_TABR(Cpt_M)
    
    Cpt_P = Cpt_P + 4
Next Cpt_M
        
End Sub


