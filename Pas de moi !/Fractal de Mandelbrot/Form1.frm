VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fractal"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSwapWin 
      Caption         =   "Changer de sortie lors du zoom"
      Height          =   855
      Left            =   5040
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdZmImg 
      Caption         =   "<--- Zoom + Image"
      Height          =   735
      Index           =   1
      Left            =   5160
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdZmImg 
      Caption         =   "---> Zoom + Image"
      Height          =   735
      Index           =   0
      Left            =   5040
      TabIndex        =   16
      Top             =   1320
      Width           =   855
   End
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6600
      Top             =   240
   End
   Begin VB.OptionButton optPct 
      Caption         =   "Sortie de droite"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton optPct 
      Caption         =   "Sortie de gauche"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Index           =   1
      Left            =   6120
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   13
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtIter 
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Text            =   "20"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtUBnd 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Text            =   "2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtLBnd 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Text            =   "-2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtUBnd 
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   6
      Text            =   "2"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtLBnd 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Text            =   "-2"
      Top             =   120
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Index           =   0
      Left            =   120
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   1
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculer/Dessiner"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   5160
      X2              =   5880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape shpHiLi 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Height          =   4875
      Left            =   90
      Top             =   1290
      Width           =   4875
   End
   Begin VB.Label lblIter 
      Caption         =   "It�rations Maxi."
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblMid 
      Alignment       =   2  'Center
      Caption         =   "�"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblVar 
      Caption         =   "y allant de"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblMid 
      Alignment       =   2  'Center
      Caption         =   "�"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblVar 
      Caption         =   "x allant de"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public C As clsCpx  'Un nombre complexe quelconque.
Dim tblPix() As Byte    'Le tableau contenant toutes les valeurs du "nombre de Mandelbrot"
Dim bCalcul As Boolean  'Calcul en cours ou pas ?
Dim XZ(0 To 1) As Integer   'Sert � tracer un cadre de zoom en laissant la souris appuy�e sur une des sorties
Dim YZ(0 To 1) As Integer   ''Sortie' d�signe le picturebox sur lequel on dessine l'ensemble de Mandelbrot
Dim pctOtp As Integer   'Sur quelle sortie on trace ?
Dim ZoomData(0 To 1, 0 To 1, 0 To 1) As Double    'Les donn�es du zoom : OutputPctIndex, [X=0, Y=1], [LB=0, UB=1]
Dim IsPctUDed(0 To 1) As Boolean    'Permet de savoir si une Pctbox est '� jour', ie si, apr�s modification
                                    '�ventuelle du zoom, le fractal a �t� redessin�
Dim Rsz As Boolean  'Est-ce qu'un zoom est en train de se faire ?

Private Sub cmdCalc_Click() 'La proc�dur qui calcule puis dessine l'ensemble de Mandelbrot
If Not bCalcul Then 'Si aucun calcul n'�tait en cour auparavant, en commencer un nouveau
    Dim i As Integer, i2 As Integer 'Pour le boucles For
    Dim j As Integer    'idem
    Dim Iter As Integer 'Nombre d'it�rations maxi pour le calcul du "nombre de Mandelbrot"
    Dim lb(0 To 1) As Double    'Les donn�es du zoom : pour x
    Dim ub(0 To 1) As Double    'et pour y
    Dim Sti As Double   'La valeur du pas pour la partie r�elle de C
    Dim Stj As Double   'idem pour la partie imaginaire
    Dim cClr As Long    'La couleur du pixel en cours de dessin (cf plus bas)
    
    bCalcul = True  'Indiquer qu'un calcul est en crous
    cmdCalc.Caption = "Arr�ter" 'Changer l'affichage du bouton
    optPct(0).Enabled = False   'Ne plus autoriser les clics sur la s�lection de la fen�tre de sortie
    optPct(1).Enabled = False
    lb(0) = ZoomData(pctOtp, 0, 0)  'charger les donn�es du zoom
    lb(1) = ZoomData(pctOtp, 1, 0)
    ub(0) = ZoomData(pctOtp, 0, 1)
    ub(1) = ZoomData(pctOtp, 1, 1)
    'Calculer les valeurs des pas pour C
    Sti = 1 / pct(pctOtp).ScaleWidth * (ub(0) - lb(0))
    Stj = 1 / pct(pctOtp).ScaleHeight * (ub(1) - lb(1))
    'Et initialiser C au point en haut � gauche
    C.RealPart = lb(0)
    C.ImagPart = lb(1)
    
    Iter = Val(txtIter.Text)    'R�cup�rer le nombre d'it�rations maxi
    ReDim tblPix(0 To pct(pctOtp).ScaleWidth, 0 To pct(pctOtp).ScaleHeight) 'Redimensionner le tableau des "nombres de Mandelbrot" (bon vu que la taille des pct ne varie pas �a sert � rien :) mais on sait jamais)
    prg.Min = 0 'Fixer la valeur min
    prg.Max = pct(pctOtp).ScaleHeight   'puis max du ProgressBar
    prg.Value = 0   'Initialiser
    For i = 0 To pct(pctOtp).ScaleWidth 'Pour chaque pixel de la pct, horizontal
        For j = 0 To pct(pctOtp).ScaleHeight    'puis vertical,
            C.ImagPart = C.ImagPart + Stj   'incr�menter C
            tblPix(i, j) = MandelbrotNum(C, Iter)   'puis calculer la valeur du "nombre de mandelBrot"
        Next j  'etc
        'Incr�menter C, fixer sa partie imaginaire (on repart sur une nouvelle colonne)
        C.RealPart = C.RealPart + Sti
        C.ImagPart = lb(1)
        DoEvents    'Laisser la main � windows un minimum
        If Not bCalcul Then Exit For    'Si entretemps on a cliqu� sur "Arr�ter", ben on arr�te
        'Mettre � jour la progressbar
        prg.Value = i
    Next i  'etc
    
    'R�initialiser la progressbar ;
    'le max est fix� au point auquel on est arriv� lors du calcul (pour traiter le cas o� on s'arr�te avant)
    prg.Max = i - 1
    prg.Value = 0
    
    For i2 = 0 To i - 1 'Pour chaque colonne calcul�e
        For j = 0 To pct(pctOtp).ScaleHeight    'parcourir tous les pixels
            cClr = Int((1 - tblPix(i2, j) / Iter) * vbRed)  'Calculer la couleur
            pct(pctOtp).PSet (i2, j), cClr  'et fixer la couleur au pixel
        Next j  'etc
        DoEvents    'Laisser la main � windows
        prg.Value = i2  'mettre  jour la progressbar
    Next i2
    IsPctUDed(pctOtp) = True    'La pct a �t� mise � jour (au niveau du dessin)
    pct(pctOtp).Picture = pct(pctOtp).Image 'Fixer l'image. Je sais pas � quoi �a sert, mais bon...
    prg.Value = 0   'R�initialiser la progressbar. Jamais trop prudent.
    If Not bCalcul Then Exit Sub    'Si le calcul a d�j� �t� stopp�, sortir de la fonction.
    'Sinon, indiquer que le calcul est terminer (le stopper)
    cmdCalc_Click
Else    'Arr�te le calcul en cours
    bCalcul = False 'proprement dit :)
    cmdCalc.Caption = "Calculer/Dessiner"   'Changer l'affichage
    prg.Value = 0   'R�initialiser la progressbar
    optPct(0).Enabled = True    'Remettre en marche les optionbuttons
    optPct(1).Enabled = True
End If
End Sub

Private Sub cmdZmImg_Click(Index As Integer)    'Transf�rer les donn�es du zoom ainsi que l'image d'un pct � l'autre
'Si le pct n'a pas �t� dessin� apr�s avoir subi des modifications de zoom, redessiner ou annuler.
If Not IsPctUDed(Index) Then
    If MsgBox("Le fractal n'a pas �t� trac� dans la sortie d'origine." & vbCrLf & "Le tracer maintenant ou annuler l'op�ration ?", vbOKCancel, "Fractal") = vbCancel Then Exit Sub
    cmdCalc_Click
End If
'Sinon, transf�rer.
ZoomData(1 - Index, 0, 0) = ZoomData(Index, 0, 0)
ZoomData(1 - Index, 0, 1) = ZoomData(Index, 0, 1)
ZoomData(1 - Index, 1, 0) = ZoomData(Index, 1, 0)
ZoomData(1 - Index, 1, 1) = ZoomData(Index, 1, 1)
pct(1 - Index).Picture = pct(Index).Picture
End Sub

Private Sub Form_Activate()
Static NotFirstTime As Boolean  'Est-ce la premi�re fois que le formulaire s'affiche ?
Dim i

If Not NotFirstTime Then    'Oui - le *Not Not*FirstTime c'est pas joli mais on fait avec :)
    frmCalc.Show , Me   'Afficher la petite fen�tre bien sympa
    Me.Enabled = False  'Interdire les int�ractions avec ce formulaire
    cmdCalc_Click   'Calculer
    cmdZmImg_Click 0    'Transf�rer l'image
    Me.Enabled = True   'R�activer le formulaire
    Unload frmCalc  'Fermer la fen�tre bien sympa
    IsPctUDed(0) = True 'Dire que les pct sont mis � jour
    IsPctUDed(1) = True
    NotFirstTime = True 'Et c'est plus la premi�re fois que le formulaire est affich�
End If
End Sub

Private Sub Form_Load()
Set C = New clsCpx  'Dire que C est une instance de la classe clsCpx (manipulation de nombres complexes)
pctOtp = 0  'Mettre la sortie par d�faut � 0 = celle de gauche
ZoomData(0, 0, 0) = -2  'Initialiser les donn�es du zoom
ZoomData(0, 1, 0) = -2
ZoomData(1, 0, 0) = -2
ZoomData(1, 1, 0) = -2
ZoomData(0, 0, 1) = 2
ZoomData(0, 1, 1) = 2
ZoomData(1, 0, 1) = 2
ZoomData(1, 1, 1) = 2
End Sub

Private Sub optPct_Click(Index As Integer)  'Clic sur une des optionbuttons
pctOtp = Index  'Changer la valeur qui indique la sortie en cours d'utilisation
shpHiLi.Left = pct(pctOtp).Left - 30    'D�placer le rectangle bleu (de s�lection - Highlight)
UpdateTBs   'Et changer le texte pour les donn�es du zoom
End Sub

Private Sub pct_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Lorsque la souris est appuy�e,
If Button = 1 Then  'et que c'est le bouton gauche
    optPct_Click Index  'Changer la sortie en cours d'utilisation
    optPct(Index).Value = 1 'Changer aussi le optionbutton s�lectionn�
    XZ(0) = -X  'Mettre X dans la valeur absolue, le '-' (moins) signifiant que le rectangle de s�lection (blanc) n'est pas encore dessin�
    YZ(0) = -Y
    Rsz = False 'Au d�part, fixer rsz sur false
    tmrMouse.Enabled = True 'Et on enclenche le Timer pour l'affichage du rectangle blanc (500 ms)
End If
End Sub

Private Sub pct_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Lorsqu'on bouge la souris sur une des pct
If Button = 1 Then  'et que c'est toujours le bouton gauche qui est appuy�,
    'Si la souris a �t� boug�e de tr�s peu, quitter la proc�dure (le 5 est en pixels, pas en twips)
    If Abs(XZ(0) - X) < 5 And Abs(YZ(0) - Y) < 5 And Not Rsz Then Exit Sub
    If tmrMouse.Enabled = True Then tmrMouse_Timer  'Si le timer n'est pas d�clench�, le faire manuellement
    Rsz = True  'Indiquer que le redimensionnement a bien d�marr�
    pct(Index).Cls  'Effacer les lignes du rectangle,
    'puis les redessiner
    pct(Index).Line (XZ(0), YZ(0))-(X, YZ(0)), vbWhite
    pct(Index).Line (XZ(0), YZ(0))-(XZ(0), Y), vbWhite
    pct(Index).Line (X, Y)-(X, YZ(0)), vbWhite
    pct(Index).Line (X, Y)-(XZ(0), Y), vbWhite
End If
End Sub

Private Sub pct_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On lib�re un bouton de la souris
If Button = 1 Then  'si c'est le bouton gauche,
    tmrMouse.Enabled = False    'Desactiver le timer pour le redimensionnement
    
    If XZ(0) < 0 Then Exit Sub  'Si XZ(0) (la valeur de d�part) est n�gatif, c'est que aucun redimensionnement n'a �t� op�r�.
    
    Dim tmp As Double   'permet de trier les valeur (pour pas avoir un "anti-zoom", avec des valeurs invers�es)
    
    Dim ALb(0 To 1) As Double   'Les donn�es actuelles du zoom
    Dim AUb(0 To 1) As Double
    ALb(0) = ZoomData(pctOtp, 0, 0) 'que l'on charge aussit�t
    ALb(1) = ZoomData(pctOtp, 1, 0)
    AUb(0) = ZoomData(pctOtp, 0, 1)
    AUb(1) = ZoomData(pctOtp, 1, 1)
    
    XZ(1) = X   'On fixe les valeurs d'arriv�e du rectangle de zoom
    YZ(1) = Y
    
    'On trie les valeurs (XZ(0) doit �tre plus petit que XZ(1) et idem pour YZ(0) et YZ(1))
    If XZ(0) > XZ(1) Then
        tmp = XZ(0)
        XZ(0) = XZ(1)
        XZ(1) = tmp
    End If
    If YZ(0) > YZ(1) Then
        tmp = YZ(0)
        YZ(0) = YZ(1)
        YZ(1) = tmp
    End If
    
    'Si on veut interchange les sorties lors du zoom (c'est plus pratique),
    If chkSwapWin.Value = 1 Then
        'ben on le fait directement
        Call optPct_Click(1 - pctOtp)
        optPct(1 - pctOtp).Value = True
    End If
    
    'On sauvegarde les nouvelles donn�es du zoom � partir des coordonn�es des points choisis
    ZoomData(pctOtp, 0, 0) = CStr(ALb(0) + (XZ(0) / pct(pctOtp).ScaleWidth) * (AUb(0) - ALb(0)))
    ZoomData(pctOtp, 0, 1) = CStr(ALb(0) + (XZ(1) / pct(pctOtp).ScaleWidth) * (AUb(0) - ALb(0)))
    ZoomData(pctOtp, 1, 0) = CStr(ALb(1) + (YZ(0) / pct(pctOtp).ScaleHeight) * (AUb(1) - ALb(1)))
    ZoomData(pctOtp, 1, 1) = CStr(ALb(1) + (YZ(1) / pct(pctOtp).ScaleHeight) * (AUb(1) - ALb(1)))
    
    'Invalider la sortie, car pour l'instant la pct n'a pas �t� redessin�e
    IsPctUDed(pctOtp) = False
    
    'Et modifier les textboxes
    UpdateTBs
Else
    'Si c'est le bouton droit, faire un zoom 'arri�re' de coefficient 2
    'on peut mettre n'importe quoi � la place de 2, bien entendu
    txtLBnd(0).Text = CStr(Val(txtLBnd(0).Text) * 2)
    txtLBnd(1).Text = CStr(Val(txtLBnd(1).Text) * 2)
    txtUBnd(0).Text = CStr(Val(txtUBnd(0).Text) * 2)
    txtUBnd(1).Text = CStr(Val(txtUBnd(1).Text) * 2)
End If
'Permet de mettre en forme correctement les textboxes
txtLBnd_LostFocus 0
txtLBnd_LostFocus 1
txtUBnd_LostFocus 0
txtUBnd_LostFocus 1
End Sub

Private Sub tmrMouse_Timer()
'Les 500ms de redimensionnement ont pass�es, valider celui-ci
'rendre les valeur positives
XZ(0) = -XZ(0)
YZ(0) = -YZ(0)
'et arr�ter le timer
tmrMouse.Enabled = False
End Sub

Private Sub UpdateTBs()
'Mettre � jour les valeurs des bo�tes de textes � partir des donn�es du zoom
txtLBnd(0).Text = Str$(ZoomData(pctOtp, 0, 0))
txtLBnd(1).Text = Str$(ZoomData(pctOtp, 1, 0))
txtUBnd(0).Text = Str$(ZoomData(pctOtp, 0, 1))
txtUBnd(1).Text = Str$(ZoomData(pctOtp, 1, 1))
End Sub

'Lorsque l'une des quatre textboxes pert le focus, sauvegarder la nouvelle valeur du zoom.
'noter que la s�paration unit�s/dixi�mes se note � l'anglaise ('.', pas ',')

Private Sub txtLBnd_LostFocus(Index As Integer)
txtLBnd(Index).Text = Str$(Val(Replace(txtLBnd(Index).Text, ",", ".")))
ZoomData(pctOtp, Index, 0) = Val(txtLBnd(Index).Text)
End Sub

Private Sub txtUBnd_LostFocus(Index As Integer)
txtUBnd(Index).Text = Str$(Val(Replace(txtUBnd(Index).Text, ",", ".")))
ZoomData(pctOtp, Index, 1) = Val(txtUBnd(Index).Text)
End Sub

