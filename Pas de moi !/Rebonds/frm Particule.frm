VERSION 5.00
Begin VB.Form frmSimulation 
   Caption         =   "Simulation du mouvement d'une particule ( concernant les rebonds)"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtlngQ 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7440
      TabIndex        =   16
      Text            =   "10"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkModeComete 
      Caption         =   "Mode ""Comète"""
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Timer tmpRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8640
      Top             =   8760
   End
   Begin VB.HScrollBar sclVitesse 
      Height          =   375
      Left            =   6360
      Max             =   255
      Min             =   1
      TabIndex        =   11
      Top             =   7920
      Value           =   1
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Effacer le plan"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   7440
      Width           =   3495
   End
   Begin VB.CommandButton cmdConfigPart 
      Caption         =   "Configurer particule"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenererTrait 
      Caption         =   "Générer Traits"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdGererTrait 
      Caption         =   "Gérer les traits"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtTemps 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6645
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "5"
      Top             =   8880
      Width           =   735
   End
   Begin VB.CheckBox optTrace 
      Caption         =   "Activer/Désactiver trace de la particule"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   8400
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Pause / Marche"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdPasAPAs 
      Caption         =   "Pas a Pas"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   8040
      Top             =   8760
   End
   Begin VB.PictureBox Plan 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label lblY 
      Caption         =   "Y ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label lblX 
      Caption         =   "X ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label lbllngQ 
      Caption         =   "Longueur de la queue :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblTemps 
      AutoSize        =   -1  'True
      Caption         =   "Interval de temps :                   ms"
      Height          =   195
      Left            =   5280
      TabIndex        =   14
      Top             =   8880
      Width           =   2355
   End
   Begin VB.Label lblK 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   12
      Top             =   8000
      Width           =   375
   End
   Begin VB.Label lblVitesse 
      Caption         =   "Vitesse (Vi=K .Vo), K="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   10
      Top             =   7995
      Width           =   2355
   End
End
Attribute VB_Name = "frmSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************'
'*************************************************************'
'****                                                     ****'
'****          Zeroc00l 2003  SimulPart v2.xx             ****'
'****                                                     ****'
'****            Détermination d'une trajectoire.         ****'
'****                                                     ****'
'*************************************************************'
'*************************************************************'

'Particule
Public X As Single, Y As Single 'Coordonées : abscisse et ordonné
Public U As Single, V As Single 'Vecteur horzontal et vertical
Dim VK As Byte
'Traits
Public nbrTrait 'Nombre de Trait
Public OptBord As Boolean ' Ls bords sont il des trait ?

Dim LongQ As Byte 'longueur de trajectoire garde ...

Private Sub Form_Initialize()
 LongQ = 10
 ReDim lstQueue(LongQ)
 Dim z As Byte
 For z = 0 To LongQ
     ReDim lstQueue(z).lstPoint(1)
     lstQueue(z).lstPoint(1).X = -1
     lstQueue(z).lstPoint(1).Y = -1
 Next z
 
 OptBord = True
 nbrTrait = 4
 VK = 10
 sclVitesse.Value = VK
 lblK.Caption = VK
 Plan.Scale (0, Plan.ScaleHeight - 1)-(Plan.ScaleWidth, -1)
End Sub

Public Sub Form_Load()
 Randomize
 Call CreerBord(nbrTrait)
End Sub


Private Sub cmdGenererTrait_Click()
 If nbrTrait = 255 Then
    MsgBox "Trop de traits..." & vbCrLf & "Vous y voyez encore quelquechose pour vouloir en rajouter ?"
 Else
  Dim z As Byte
  z = nbrTrait
  nbrTrait = Int(10 * Rnd) + 4 + nbrTrait
  If nbrTrait > 255 Then nbrTrait = 255
  ReDim Preserve lstT(nbrTrait)
  
  If frmSimulation.OptBord = True Then
     nbrTrait = nbrTrait - z
     For z = z - 3 To z
         lstT(z + nbrTrait).M = lstT(z).M
         lstT(z + nbrTrait).N = lstT(z).N
         lstT(z + nbrTrait).Y1 = lstT(z).Y1
         lstT(z + nbrTrait).Y2 = lstT(z).Y2
         lstT(z + nbrTrait).a = lstT(z).a
         lstT(z + nbrTrait).b = lstT(z).b
     Next z
  End If
  Dim Var1 As Integer, Var2 As Integer
  For z = z - 4 To z + nbrTrait - 1 + 4 * (OptBord = True)
   Var1 = Int(601 * Rnd)
   Var2 = Int(601 * Rnd)
   lstT(z).M = IIf(Var1 <= Var2, Var1, Var2)
   lstT(z).N = IIf(Var2 >= Var1, Var2, Var1)
   lstT(z).Y1 = Int(481 * Rnd)
   lstT(z).Y2 = Int(481 * Rnd)
   If lstT(z).N <> lstT(z).M Then
      lstT(z).a = (lstT(z).Y2 - lstT(z).Y1) / (lstT(z).N - lstT(z).M)
      lstT(z).b = lstT(z).Y1 - lstT(z).a * lstT(z).M
   End If
  Next z
  nbrTrait = z - 4 * (OptBord = True) - 1
  
  Call DessinerTrait(1)
 End If
End Sub

Private Sub cmdGererTrait_Click()
 frmGestionTrait.Show
End Sub
Private Sub cmdConfigPart_Click()
 frmConfigPart.Show
End Sub
Private Sub cmdClear_Click()
 Plan.Cls
 Call DessinerTrait(1)
End Sub


Private Sub cmdGo_Click()

'Création des traits

'nbrTrait = 3
'ReDim lstT(nbrTrait)
'lstT(1).M = 600
'lstT(1).N = 600
'lstT(2).M = 0
'lstT(2).N = 600
'lstT(3).M = 0
'lstT(3).N = 600
'
'lstT(1).a = 0
'lstT(1).b = 0
'lstT(2).a = -240 / 600
'lstT(2).b = 240
'lstT(3).a = 240 / 600
'lstT(3).b = 240
'
'For Z = 1 To nbrTrait
' Plan.Line (lstT(Z).M, lstT(Z).a * lstT(Z).M + lstT(Z).b)-(lstT(Z).N, lstT(Z).a * lstT(Z).N + lstT(Z).b)
'Next Z
'Plan.Line (600, 0)-(600, 480), vbRed

 
 
'Initialisation

'Demarrage au centre : Coordonnées...
'Vecteurs...

If X = 0 And Y = 0 Then
   X = Int(601 * Rnd)
   Y = Int(481 * Rnd)
End If
If U = 0 And V = 0 Then
   Do
     U = Int(20 * Rnd) - 12
   Loop Until U
   Do
     V = Int(20 * Rnd) - 12
   Loop Until V
End If

Timer.Interval = Val(txtTemps.Text)
Timer.Enabled = True
tmpRefresh.Enabled = True

End Sub

'Info sur le point visé

Private Sub Plan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblX.Caption = "X = " & X
 lblY.Caption = "Y = " & Y
End Sub



Private Sub cmdStop_Click()
 Timer.Enabled = Not (Timer.Enabled)
 tmpRefresh.Enabled = Not (tmpRefresh.Enabled)
End Sub

Private Sub cmdPasAPAs_Click()
 Call Timer_Timer
 lblX.Caption = X
 lblY.Caption = Y


If optTrace.Value = True Then Plan.PSet (Ax, By), vbWhite
 Plan.PSet (X, Y)
End Sub

Private Sub sclVitesse_Change()
 U = U * sclVitesse.Value / VK
 V = V * sclVitesse.Value / VK
 VK = sclVitesse.Value
 lblK.Caption = sclVitesse.Value
End Sub

Private Sub chkModeComete_Click()
 lbllngQ.Visible = Not (lbllngQ.Visible)
 txtlngQ.Visible = Not (txtlngQ.Visible)
 tmpRefresh.Enabled = Not (tmpRefresh.Enabled)
 Call DessinerTrait(1)
End Sub

Private Sub txtlngQ_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtlngQ_Change()
 If Val(txtlngQ.Text) > 255 Then txtlngQ.Text = "255"
 If Val(txtlngQ.Text) < 2 Then txtlngQ.Text = "2"
  ReDim Preserve lstQueue(Val(txtlngQ.Text))
 Dim z As Byte
'On itnialise les nouveaux segments de queue
'Je suis obligé de mettre val(...)-1 car silngQ=255, la boucle attribue la valeur 256 a z
 If LongQ < Val(txtlngQ.Text) Then 'sans ce test longQ peut etre = a 255
  For z = LongQ + 1 To Val(txtlngQ.Text) - 1 'Et la z =255+1 ... erreur de dépassement
      ReDim lstQueue(z).lstPoint(1)
      lstQueue(z).lstPoint(1).X = -1
      lstQueue(z).lstPoint(1).Y = -1
  Next z
  ReDim lstQueue(z).lstPoint(1)
  lstQueue(z).lstPoint(1).X = -1
  lstQueue(z).lstPoint(1).Y = -1
 End If
 LongQ = Val(txtlngQ.Text)
End Sub

Private Sub txtTemps_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtTemps_Change()
 Timer.Interval = Val(txtTemps.Text)
End Sub


Private Sub tmpRefresh_Timer()
 Call DessinerTrait(0)
End Sub

'***********************************************************
'*                  Le Code Moteur ....                    *
'***********************************************************

Private Sub Timer_Timer()
 
If chkModeComete.Value = Checked Then
 Static Pointeur As Byte 'Pointe vers un tronçon de queue
 If Pointeur + 1 > LongQ Then Pointeur = 1 Else Pointeur = Pointeur + 1
 Pl = 1 '1ère parti du tronçon
 lstQueue(Pointeur).lstPoint(Pl).X = X 'Si la trajectoire se
 lstQueue(Pointeur).lstPoint(Pl).Y = Y 'décompose en plusieurs parties...
End If
 
 
 
 Dim i As Single, P As Byte, K As Single 'Variables...
 
 'Variables de sauvegarde...
 Dim MR As Single, TC As Byte 'Meilleur Rapport t Trait correspondant
 Dim DTT As Byte 'Dernier Trait Touché
 Dim DR As Single 'Dernier Rapport si la particule est réfléchie !
 Dim PPR As Single ' Plus petit Rapport permis
 
 'Nota : ICI : DTT = 0     '
             ' DR  = 0     '' N'impose aucune contrainte dans le premier teste qui suit
             ' PPR = 0     '
 Do

   
   TC = 0 'Par défaut aucun trait ne correspond
   MR = 1 'Par défaut il n'existe pas de meilleur rapport. On met donc un nbr>=1
   P = 0
   
   'Obtention du trait réflecteur ( si y'en a un ) ou Trait Correspondant (TC)
   Do Until P = nbrTrait Or P = nbrTrait - 1 And DTT = nbrTrait
      K = 10 'On fixe un rapport >1

      P = P + 1 - (P + 1 = DTT) 'on ne reteste pas le trait qui vient de Réflechir
      
      'K retourne la portion du vecteur pour entrer en collision (le vecteur ne tombe pas pile sur le trait, il le coupe )
      ' Plus c'est petit, mieux c'est. Si k> a 1 on ne s'y interesse qu'au prochain mouvement.
      'I : la coordonné X de la colision
      
      
      If lstT(P).M <> lstT(P).N Then 'Cas des trait non verticaux
         'On vérifie que le vecteur n'est pas colinéaire au trait ... sinon k-> infini et I n'existe pas
         If U * lstT(P).a - V <> 0 Then  'Le cas contraire signifie que le vecteur (U,V) est colineaire au trait
            i = (U * (Y - lstT(P).b) - V * X) / (U * lstT(P).a - V)
            K = (Y - lstT(P).b - lstT(P).a * X) / (lstT(P).a * U - V)
         End If
      Else 'Case des trait verticaux
         If Abs(lstT(P).M - X) <= Abs(U) Then i = lstT(P).M
         If U <> 0 Then K = (lstT(P).M - X) / U
      End If
      
      'Choisis le meilleurs trait candidat à la reflexion
      ' Le k trait doit etre inferieur a 1 sinon pas de collision
      ' il doit etre meilleur que celui deja existant donc inferieur a MR
      ' il doit être supérieur (ou égale) au Plus Petit Rapport Permis K>= PPR
      If K > 0 And K <= 1 And i >= lstT(P).M And i <= lstT(P).N _
               And K <= MR And K >= PPR Then
         'Enregistrment du trait dont on devra se servir pour le calcul de X,Y et U,V
         MR = K 'le Meilleur Rapport (MR) deviendra le PPR du test qui suit
         TC = P 'Trait correspondant, le plus proche
      End If
      
      
    'End If
   Loop
      
   'option de traçage
   If chkModeComete.Value = Checked Then
     If Pl + 1 < 2 ^ 15 Then
        Pl = Pl + 1
     Else
        MsgBox " Erreur, trop de rebond de la particule." & vbCrLf & "Veuillez diminuer le facteur K ou bien dessiner un schéma de trait"
        Call cmdStop_Click
        Exit Sub
     End If
     ReDim Preserve lstQueue(Pointeur).lstPoint(Pl)
     lstQueue(Pointeur).lstPoint(Pl).X = X + MR * U
     lstQueue(Pointeur).lstPoint(Pl).Y = Y + MR * V
   ElseIf optTrace.Value = Checked Then
     Plan.Line (X + PPR * U, Y + PPR * V)-(X + MR * U, Y + MR * V), vbGreen
   Else
     Plan.PSet (Ax, By), vbWhite
   End If
   
   'Si il existe un TC alors il faut recalculer
   'sinon, c'est fini X=X+U et Y = Y+U
   If TC Then
     'Si droite non vertical
     If lstT(TC).M <> lstT(TC).N Then
        i = (-U * lstT(TC).a ^ 2 + 2 * lstT(TC).a * V + U) / (lstT(TC).a ^ 2 + 1)
        V = (V * lstT(TC).a ^ 2 + 2 * lstT(TC).a * U - V) / (lstT(TC).a ^ 2 + 1)
        U = i
        i = (lstT(TC).a * (2 * (Y - lstT(TC).b) - lstT(TC).a * X) + X) / (lstT(TC).a ^ 2 + 1)
        Y = (2 * (lstT(TC).a * X + lstT(TC).b) + Y * (lstT(TC).a ^ 2 - 1)) / (lstT(TC).a ^ 2 + 1)
        X = i
     Else 'droite vertical
        X = 2 * lstT(TC).M - X
        U = -U
     End If
            
     'Enregistrement du
     DTT = TC 'Dernier Trait Touché
     PPR = MR  'Rapport minimum

   End If
 Loop Until TC = 0 'Tant qu'il y a des collision, on reteste
 
 'Plus de collision, on calcule simplement les nouvelles coordonnées
 X = X + U
 Y = Y + V
 
 'Affichage des Infos
 lblX.Caption = "X = " & Round(X, 2) 'Coord X de la particule
 lblY.Caption = "Y = " & Round(Y, 2) 'Coord Y
 Plan.PSet (X, Y), vbRed
 
  If chkModeComete.Value = Checked Then lstQueue(Pointeur).lstPoint(0).X = Pl
 If chkModeComete.Value Then Call DrawComete(LongQ, Pointeur)
 End Sub
 
