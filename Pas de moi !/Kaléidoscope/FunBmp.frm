VERSION 5.00
Begin VB.Form FunBmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kaléidoscope "
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3090
   FillStyle       =   0  'Solid
   Icon            =   "FunBmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   3090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PB_RECH 
      Caption         =   "Réinit"
      Height          =   300
      Left            =   2295
      TabIndex        =   15
      Top             =   4065
      Width           =   705
   End
   Begin VB.CommandButton PB_SAV 
      Caption         =   "Sauve"
      Height          =   300
      Left            =   2280
      TabIndex        =   14
      Top             =   3450
      Width           =   750
   End
   Begin VB.CommandButton PB_Stop 
      Caption         =   "Stop"
      Height          =   300
      Left            =   2295
      TabIndex        =   13
      Top             =   3135
      Width           =   750
   End
   Begin VB.CommandButton PB_QQQ 
      Caption         =   "?"
      Height          =   300
      Left            =   1695
      TabIndex        =   12
      Top             =   3615
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      Height          =   3060
      Left            =   15
      Picture         =   "FunBmp.frx":0CCA
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   10
      Top             =   15
      Width           =   3060
   End
   Begin VB.Timer Timer1 
      Interval        =   17
      Left            =   1650
      Top             =   3525
   End
   Begin VB.CommandButton Command3 
      Caption         =   "decalG"
      Height          =   300
      Index           =   9
      Left            =   780
      TabIndex        =   9
      Top             =   3135
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Miroir 3"
      Height          =   300
      Index           =   8
      Left            =   780
      TabIndex        =   8
      Top             =   3750
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Miroir 2"
      Height          =   300
      Index           =   7
      Left            =   30
      TabIndex        =   7
      Top             =   3750
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Miroir 1"
      Height          =   300
      Index           =   6
      Left            =   1530
      TabIndex        =   6
      Top             =   4065
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DecalD"
      Height          =   300
      Index           =   5
      Left            =   1530
      TabIndex        =   5
      Top             =   3135
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pitha 2"
      Height          =   300
      Index           =   4
      Left            =   780
      TabIndex        =   4
      Top             =   4065
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clair(+)"
      Height          =   300
      Index           =   3
      Left            =   30
      TabIndex        =   3
      Top             =   3435
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clair(-)"
      Height          =   300
      Index           =   2
      Left            =   780
      TabIndex        =   2
      Top             =   3435
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pitha 1"
      Height          =   300
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   4065
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fibo"
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   3135
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Supporte le drag and drop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   315
      TabIndex        =   11
      Top             =   4440
      Width           =   2265
   End
End
Attribute VB_Name = "FunBmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'* Bouton générique
Private Sub Command3_Click(Index As Integer)
Dim L_R As Integer
Dim L_V As Integer
Dim L_B As Integer
Dim L_Chaine As String
Dim L_String_Col As String
Dim L_Comp_Col As String
Dim L_stat_Col As String
Dim L_lastbytvalide As Long
Dim L_axe1 As Long
Dim L_axe2 As Long
Dim k As Long
Dim L As Long
Dim Decal As Long
Dim v_Decal As Long
Dim L_test As Integer
Dim L_TmpByte As Byte
Dim L_Larg As Long
Dim L_Haut As Long
Dim L_PicSav As Byte
Dim L_SavByte As M_RVB
Dim L_Compt As Long
Dim T_Ligne As Long
Dim L_Deb As Long
Dim L_debTmp As Long
Dim L_Fin As Long
Dim L_stop As Boolean
Dim Fin As Long
Dim Boucle As Long
Dim L_sav As String

On Error Resume Next

k = 1
L_Chaine = ""
L_String_Col = ""
L_Comp_Col = ""
L_test = 0
L_stat_Col = ""
L_lastbytvalide = 0
L_PicSav = 0
L_Compt = 0
T_Ligne = True
L_sav = Me.Label1.Caption
Me.Label1.Caption = "traitement en cours"
Me.Label1.Refresh

G_Pause = True

'* ou par l'application (ici par l'application propriété picture!!!)
GetObject Picture1.Image, Len(PicInfo), PicInfo
'* surcharge perso car le mode de calcul n'est pas juste
L_lastbytvalide = PicInfo.bmWidth * PicInfo.bmHeight * 4
L_Larg = PicInfo.bmWidth
L_Haut = PicInfo.bmHeight
        
'Copy the bitmapbits to the array
GetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
    
L_Chaine = ""

Select Case Index
    
    '* décalage ligne + 1 sur chaque nouvelle ligne
    Case 9
        k = 1
        Cnt = 1
        '* chargement en mémoire dans les tableaux de travail
        '* on fait l'économie d'un tableau car image en 24 bits !!! et non 32
        For Cnt = 1 To L_lastbytvalide Step 4
            G_TABB(k) = PicBits(Cnt)
            G_TABV(k) = PicBits(Cnt + 1)
            G_TABR(k) = PicBits(Cnt + 2)
            
            k = k + 1
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
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
                 
                '* Gestion des indices de débordement de pile
                If v_Decal > L_Fin Then
                    v_Decal = L_debTmp
                    L_debTmp = L_debTmp + 1
                End If
                
                PicBits(k) = G_TABB(v_Decal)
                PicBits(k + 1) = G_TABV(v_Decal)
                PicBits(k + 2) = G_TABR(v_Decal)
                
                '* gestion des indices du tableau en 32 bits
                k = k + 4
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit For
                End If
            Next L_Deb
                        
            '* gestion de la ligne
            L = L_Deb
            L_debTmp = L_Deb - 1
            L_Fin = L_Fin + L_Larg
            
            '* gestion du décalage toujour inférieur à la largeur du tableau
            If Decal = (L_Larg - 1) Then
                Decal = 1
            Else
                Decal = (Decal) + 1
            End If
                        
        Next T_Ligne
    
    'Retournement ?
    Case 8
        '* Détermine axe de symetrie on pat de l'axe on va vers 0
        '* cas axe pair ou cas axe impair
        '* on mape l'image dans un tableau à 4 dimensions(largeur) et de la hauteur
        '* de l'image xxx
        k = 1
        For Cnt = 1 To L_lastbytvalide Step 4
            G_TABB(k) = PicBits(Cnt)
            G_TABV(k) = PicBits(Cnt + 1)
            G_TABR(k) = PicBits(Cnt + 2)
            
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            
            k = k + 1
        Next Cnt
        
        '* Axe de symétrie sur la matrice
        L_Axe = (k - 1) / 2
        
        '* mise à jour en mémmoire
        Decal = L_Axe + 1
        For Cnt = 1 To L_Axe
            '* permutation
            'on sauve la destination du premier octet R
            'permut R
            L_TmpByte = G_TABB(Decal)
            G_TABB(Decal) = G_TABB(Cnt)
            G_TABB(Cnt) = L_TmpByte
            
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            
            'PermutV
            L_TmpByte = G_TABV(Decal)
            G_TABV(Decal) = G_TABV(Cnt)
            G_TABV(Cnt) = L_TmpByte
            
            'PermutB
            L_TmpByte = G_TABR(Decal)
            G_TABR(Decal) = G_TABR(Cnt)
            G_TABR(Cnt) = L_TmpByte
            
            '* Déplacement du symetrique suivant
            Decal = (Decal) + 1
        Next Cnt
        
        '* restitiution
        k = 1
        For Cnt = 1 To (L_Axe * 2)
            PicBits(k) = G_TABB(Cnt)
            PicBits(k + 1) = G_TABV(Cnt)
            PicBits(k + 2) = G_TABR(Cnt)
            k = k + 4
        Next Cnt
        
        '* mise à jour graphique
        'SetBitmapBits Me.Picture, UBound(PicBits), PicBits(1)
        'Me.Refresh
    
    '* retournement horizontale
    Case 7
        '* Détermine axe de symetrie on part de l'axe on va vers 0
        '* cas axe pair ou cas axe impair
        '* on mape l'image dans un tableau à 4 dimensions(largeur) et de la hauteur
        '* de l'image xxx
        k = 1
        For Cnt = 1 To L_lastbytvalide Step 4
            G_TABB(k) = PicBits(Cnt)
            G_TABV(k) = PicBits(Cnt + 1)
            G_TABR(k) = PicBits(Cnt + 2)
            
            k = k + 1
        Next Cnt
        
        '* Axe de symétrie sur la matrice
        L_Axe = (k - 1) / 2
        
        '* mise à jour en mémmoire
        Decal = L_Axe + 1
        For Cnt = L_Axe To 1 Step -1
            '* permutation
            'on sauve la destination du premier octet R
            'permut R
            L_TmpByte = G_TABB(Decal)
            G_TABB(Decal) = G_TABB(Cnt)
            G_TABB(Cnt) = L_TmpByte
            
            'PermutV
            L_TmpByte = G_TABV(Decal)
            G_TABV(Decal) = G_TABV(Cnt)
            G_TABV(Cnt) = L_TmpByte
            
            'PermutB
            L_TmpByte = G_TABR(Decal)
            G_TABR(Decal) = G_TABR(Cnt)
            G_TABR(Cnt) = L_TmpByte
            
            '* Déplacement du symetrique suivant
            Decal = (Decal) + 1
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next Cnt
        
        '* restitiution
        k = 1
        For Cnt = 1 To (L_Axe * 2)
            PicBits(k) = G_TABB(Cnt)
            PicBits(k + 1) = G_TABV(Cnt)
            PicBits(k + 2) = G_TABR(Cnt)
            k = k + 4
        Next Cnt
                
    Case 6
        'ReDim G_TABR(1 To L_lastbytvalide) As Byte
        'ReDim G_TABV(1 To L_lastbytvalide) As Byte
        'ReDim G_TABB(1 To L_lastbytvalide) As Byte
        '* Détermine axe de symetrie on pat de l'axe on va vers 0
        '* cas axe pair ou cas axe impair
        '* on mape l'image dans un tableau à 4 dimensions(largeur) et de la hauteur
        '* de l'image xxx
        k = 1
        For Cnt = 1 To L_lastbytvalide Step 4
            G_TABB(k) = PicBits(Cnt)
            G_TABV(k) = PicBits(Cnt + 1)
            G_TABR(k) = PicBits(Cnt + 2)
            
            k = k + 1
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next Cnt
        
        '* Axe de symétrie sur la matrice
        L_Axe = (k - 1) / 2
        
        '* mise à jour en mémmoire
        Decal = L_Axe + 1
        For Cnt = L_Axe To 1 Step -1
            '* permutation
            'on sauve la destination du premier octet R
            'permut R
            L_TmpByte = G_TABB(Cnt)
            G_TABB(Decal) = G_TABB(Cnt)
            G_TABB(Cnt) = L_TmpByte
            
            'PermutV
            L_TmpByte = G_TABV(Cnt)
            G_TABV(Decal) = G_TABV(Cnt)
            G_TABV(Cnt) = L_TmpByte
            
            'PermutB
            L_TmpByte = G_TABR(Cnt)
            G_TABR(Decal) = G_TABR(Cnt)
            G_TABR(Cnt) = L_TmpByte
            
            '* Déplacement du symetrique suivant
            Decal = (Decal) + 1
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next Cnt
        
        '* restitiution
        k = 1
        For Cnt = 1 To (L_Axe * 2)
            PicBits(k) = G_TABB(Cnt)
            PicBits(k + 1) = G_TABV(Cnt)
            PicBits(k + 2) = G_TABR(Cnt)
            k = k + 4
        Next Cnt
               
    Case Else
        '* modification pour matrisation RVB normalement le step est de 1
        '* et chaque valeur RVB est inversée une via la boucle
        For Cnt = 1 To L_lastbytvalide Step 4
            'UBound(PicBits) Step 4
            If Cnt + 3 <= L_lastbytvalide Then
            'If Cnt + 3 <= UBound(PicBits) Then
        
                '* mise à blanc des valeurs exclut du périmetre
                '* Application d'un filtre
                If Cnt > 9 _
                And (Index = 0 Or Index = 1) Then
                    '************************  Bonni ******************************
                    If Index = 0 Then
                        PicBits(Cnt) = MonModBonifacio(PicBits(Cnt - 4), PicBits(Cnt - 8))
                        PicBits(Cnt + 1) = MonModBonifacio(PicBits(Cnt - 3), PicBits(Cnt - 7))
                        PicBits(Cnt + 2) = MonModBonifacio(PicBits(Cnt - 2), PicBits(Cnt - 6))
                
                    '************************  Pytha  ******************************
                    '* Pytha 1 RVB
                    ElseIf Index = 1 Then
                        PicBits(Cnt) = Monpythagore(PicBits(Cnt - 4), PicBits(Cnt - 8))
                        PicBits(Cnt + 1) = Monpythagore(PicBits(Cnt - 3), PicBits(Cnt - 7))
                        PicBits(Cnt + 2) = Monpythagore(PicBits(Cnt - 2), PicBits(Cnt - 6))
                    End If
        
                '* decalage
                Else
                    '* Assombri : Decal -
                    If Index = 2 Then
                        PicBits(Cnt) = Mondecalage(PicBits(Cnt), 10)
                        PicBits(Cnt + 1) = Mondecalage(PicBits(Cnt + 1), 10)
                        PicBits(Cnt + 2) = Mondecalage(PicBits(Cnt + 2), 10)
            
                    '* Eclairci : decal +
                    ElseIf Index = 3 Then
                        PicBits(Cnt) = Mondecalage(PicBits(Cnt), -10)
                        PicBits(Cnt + 1) = Mondecalage(PicBits(Cnt + 1), -10)
                        PicBits(Cnt + 2) = Mondecalage(PicBits(Cnt + 2), -10)
                            
                    '* Pytha 2 RV -> b, Br-> V, VB ->
                    ElseIf Index = 4 Then
                        If Cnt = 1 Then
                            PicBits(Cnt + 2) = Monpythagore(PicBits(Cnt), PicBits(Cnt + 1))
                        Else
                            PicBits(Cnt) = Monpythagore(PicBits(Cnt - 3), PicBits(Cnt - 2))
                            PicBits(Cnt + 1) = Monpythagore(PicBits(Cnt - 2), PicBits(Cnt))
                            PicBits(Cnt + 2) = Monpythagore(PicBits(Cnt), PicBits(Cnt + 1))
                        End If
            
                    '* décalage des pixels sur la droite
                    ElseIf Index = 5 Then
                
                        'gestion des indices
                        If Cnt = 1 Then
                            Decal = (4 * 12) + 1
                            k = Decal
                        Else
                            k = k + 4
                        End If
                
                        If k >= UBound(PicBits) Then
                            k = 1
                        End If
                                                    
                        '* Maj du pixel (valeurs RVB)
                        PicBits2(Cnt) = PicBits(k)
                        PicBits2(Cnt + 1) = PicBits(k + 1)
                        PicBits2(Cnt + 2) = PicBits(k + 2)
                    End If
                End If
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit For
                End If
            End If
            
        Next Cnt
End Select
    
'* affectation memoire
SetBitmapBits Me.Picture1.Image, UBound(PicBits), PicBits(1)

'* Gestion graphique
Me.Label1.Caption = L_sav
Me.Label1.Refresh

'* gestion du multithreading
If Me.PB_Stop.Caption = "Go" Then
    Me.Picture1.Refresh
    G_Pause = True
Else
    Me.PB_Stop.Caption = "Stop"
    Me.PB_SAV.Visible = False
    G_Pause = False
End If

End Sub


Private Sub Form_Load()
    '* gestion mémoire
    Call Init(Me)
    
    G_Time = 30
    G_Pas = -2
        
    '* Gestion graphique
    Me.Picture1.OLEDragMode = 0
    Me.Picture1.OLEDropMode = 1
    
    Me.Command3(5).Visible = False
    Me.Command3(6).Visible = False
    Me.PB_SAV.Visible = False
    
    Me.Picture1.ToolTipText = "(Utiliser le drag and drop pour changer l'image...Use drag and drop to insert a picture...)"
End Sub


Private Sub PB_QQQ_Click()
    MsgBox "créer le 13/dec/2004 par Vincent Dallaporta", , "For The Fun"
End Sub



Private Sub PB_SAV_Click()
Dim L_Str As String
    
    L_Str = InputBox("Taper le nom" & Chr(13) & "(type the name)", "Yes....", "example")
    
    L_Str = App.Path & "\" & L_Str & ".bmp"
    SavePicture Me.Picture1.Image, L_Str
End Sub

Private Sub PB_Stop_Click()
    If Me.PB_Stop.Caption = "Stop" Then
        G_Pause = True
        Me.PB_Stop.Caption = "Go"
        Me.PB_SAV.Visible = True
    Else
        Me.PB_Stop.Caption = "Stop"
        G_Pause = False
        Me.PB_SAV.Visible = False
    End If
    
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim L_data As Variant
Dim i As Double

On Error Resume Next
If Data.GetFormat(15) Then
    '* Chargement et ajustement de la fenêtre
    Me.Picture1 = LoadPicture(Data.Files.Item(1))

    If Err.Number <> 0 Then
        MsgBox "Probleme sur le fichier)" & Chr(13) & fn, , "Alerte"
        Err.Clear
    End If
    Me.Picture1.Refresh
End If

'* ou par l'application (ici par l'application propriété picture!!!)
GetObject Me.Picture1.Image, Len(PicInfo), PicInfo
G_Large = PicInfo.bmWidth
G_Haut = PicInfo.bmHeight
    
'* surcharge perso car le mode de calcul n'est pas juste
ReDim PicBits(1 To (PicInfo.bmWidth * PicInfo.bmHeight * 4)) As Byte
        
'* Copy the bitmapbits to the array
GetBitmapBits Me.Picture1.Image, UBound(PicBits), PicBits(1)
    
'Préparation des espaces de travail mémoire
Call REDIM_GTAB_BVR
End Sub

Private Sub Timer1_Timer()
Dim L_Larg As Long
Dim L_Haut As Long
Dim hIcon As Long, hDuplIcon As Long
Dim Pas As Integer

If G_Pause = False Then
    'G_Time = 5000
    Me.Timer1.Interval = G_Time
    
    '* Remplissage avec séparation des couches
    Call Dematrisation_GTAB_BVR
                
    '* Lancement du traitement de rotation
    Call Rotation1PixSurAxe(G_Large, G_Haut)

    '* Restitution Mémoire
    Call RESTIT_MEM
        
    '* mise à jour graphique
    SetBitmapBits Me.Picture1.Image, UBound(PicBits), PicBits(1)
    Me.Picture1.Refresh
    
    Select Case G_Time
        Case 18
            G_Pas = 2
        Case 36
            G_Pas = -2
    End Select
    G_Time = G_Time + G_Pas
End If
         
End Sub



Private Sub PB_RECH_Click()
    '* RéAffectation mémoire de l'image d'origine
    SetBitmapBits Me.Picture1.Image, UBound(PicBitsT), PicBitsT(1)
    GetBitmapBits Me.Picture1.Image, UBound(PicBits), PicBits(1)
    Me.Picture1.Refresh
End Sub

    


