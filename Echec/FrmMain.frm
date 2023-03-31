VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Echec"
   ClientHeight    =   6330
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":0000
   ScaleHeight     =   6330
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LblJoue 
      Alignment       =   2  'Center
      Caption         =   "Blanc Joue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Shape ShpSel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   735
      Left            =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   7
      Left            =   5040
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   6
      Left            =   4320
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   5
      Left            =   3600
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   4
      Left            =   2880
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   3
      Left            =   2160
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   2
      Left            =   1440
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   1
      Left            =   720
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   0
      Left            =   0
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   63
      Left            =   5040
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   62
      Left            =   4320
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   61
      Left            =   3600
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   60
      Left            =   2880
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   59
      Left            =   2160
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   58
      Left            =   1440
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   57
      Left            =   720
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   56
      Left            =   0
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   55
      Left            =   5040
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   54
      Left            =   4320
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   53
      Left            =   3600
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   52
      Left            =   2880
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   51
      Left            =   2160
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   50
      Left            =   1440
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   49
      Left            =   720
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   48
      Left            =   0
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   47
      Left            =   5040
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   46
      Left            =   4320
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   45
      Left            =   3600
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   44
      Left            =   2880
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   43
      Left            =   2160
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   42
      Left            =   1440
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   41
      Left            =   720
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   40
      Left            =   0
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   39
      Left            =   5040
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   38
      Left            =   4320
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   37
      Left            =   3600
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   36
      Left            =   2880
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   35
      Left            =   2160
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   34
      Left            =   1440
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   33
      Left            =   720
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   32
      Left            =   0
      Top             =   3360
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   31
      Left            =   5040
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   30
      Left            =   4320
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   29
      Left            =   3600
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   28
      Left            =   2880
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   27
      Left            =   2160
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   26
      Left            =   1440
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   25
      Left            =   720
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   24
      Left            =   0
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   23
      Left            =   5040
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   22
      Left            =   4320
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   21
      Left            =   3600
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   20
      Left            =   2880
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   19
      Left            =   2160
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   18
      Left            =   1440
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   17
      Left            =   720
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   16
      Left            =   0
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   15
      Left            =   5040
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   14
      Left            =   4320
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   13
      Left            =   3600
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   12
      Left            =   2880
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   11
      Left            =   2160
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   10
      Left            =   1440
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   9
      Left            =   720
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   8
      Left            =   0
      Top             =   1200
      Width           =   765
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu MenuFichierNew 
         Caption         =   "Nouvelle Partie"
      End
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tEchiquier(63) As String
Dim vSel(2) As String ' 0: Nom de la pièce  1: Numéro de la case   2: Couleur de la case
Dim tInfoPion(1) As String ' 0: Couleur de la pièce  1: Nom de la pièce
Dim vOldCouleur As String

Private Sub Form_Load()
Dim vCount As Integer
Dim vFlip As Integer
    
    tEchiquier(0) = "TourNoir"
    tEchiquier(1) = "ChevalNoir"
    tEchiquier(2) = "FouNoir"
    tEchiquier(3) = "ReineNoir"
    tEchiquier(4) = "RoiNoir"
    tEchiquier(5) = "FouNoir"
    tEchiquier(6) = "ChevalNoir"
    tEchiquier(7) = "TourNoir"
    For vCount = 8 To 15
        tEchiquier(vCount) = "PionNoir"
    Next
    For vCount = 16 To 47
        tEchiquier(vCount) = ""
    Next
    For vCount = 48 To 55
        tEchiquier(vCount) = "PionBlanc"
    Next
    tEchiquier(56) = "TourBlanc"
    tEchiquier(57) = "ChevalBlanc"
    tEchiquier(58) = "FouBlanc"
    tEchiquier(59) = "ReineBlanc"
    tEchiquier(60) = "RoiBlanc"
    tEchiquier(61) = "FouBlanc"
    tEchiquier(62) = "ChevalBlanc"
    tEchiquier(63) = "TourBlanc"
    
    For vCount = 0 To 63
        If vCount Mod 2 = 0 Then
            If vCount >= 0 And vCount <= 7 Then
                Image1(vCount).Picture = LoadPicture("./" & tEchiquier(vCount) & "SurBlanc.bmp")
            ElseIf vCount >= 8 And vCount <= 15 Then
                Image1(vCount).Picture = LoadPicture("./PionNoirSurNoir.bmp")
            ElseIf vCount >= 16 And vCount <= 23 Or vCount >= 32 And vCount <= 39 Then
                Image1(vCount).Picture = LoadPicture("./CaseBlanc.bmp")
            ElseIf vCount >= 24 And vCount <= 31 Or vCount >= 40 And vCount <= 47 Then
                Image1(vCount).Picture = LoadPicture("./CaseNoir.bmp")
            ElseIf vCount >= 48 And vCount <= 55 Then
                Image1(vCount).Picture = LoadPicture("./PionBlancSurBlanc.bmp")
            ElseIf vCount >= 56 And vCount <= 63 Then
                Image1(vCount).Picture = LoadPicture("./" & tEchiquier(vCount) & "SurNoir.bmp")
            End If
        Else
            If vCount >= 0 And vCount <= 7 Then
                Image1(vCount).Picture = LoadPicture("./" & tEchiquier(vCount) & "SurNoir.bmp")
            ElseIf vCount >= 8 And vCount <= 15 Then
                Image1(vCount).Picture = LoadPicture("./PionNoirSurBlanc.bmp")
            ElseIf vCount >= 16 And vCount <= 23 Or vCount >= 32 And vCount <= 39 Then
                Image1(vCount).Picture = LoadPicture("./CaseNoir.bmp")
            ElseIf vCount >= 24 And vCount <= 31 Or vCount >= 40 And vCount <= 47 Then
                Image1(vCount).Picture = LoadPicture("./CaseBlanc.bmp")
            ElseIf vCount >= 48 And vCount <= 55 Then
                Image1(vCount).Picture = LoadPicture("./PionBlancSurNoir.bmp")
            ElseIf vCount >= 56 And vCount <= 63 Then
                Image1(vCount).Picture = LoadPicture("./" & tEchiquier(vCount) & "SurBlanc.bmp")
            End If
        End If
    Next
End Sub

Private Sub Image1_Click(Index As Integer)
Dim vCount As Integer
Dim vNewCouleur As String
Static vJoue As Boolean

' 1ère sélection
    If vSel(0) = "" Then
        If tEchiquier(Index) <> "" Then
            InfoPion (Index)
            vOldCouleur = tInfoPion(0)
        End If
        If Left(tInfoPion(0), 1) = Left(LblJoue.Caption, 1) Then
            vSel(1) = Index
            vSel(0) = tEchiquier(Index)
            ShpSel.Left = Image1(Index).Left
            ShpSel.Top = Image1(Index).Top
            ShpSel.Visible = True
        End If
        
' 2ème Clique
    Else
        ShpSel.Visible = False
        If tEchiquier(Index) <> "" Then
            InfoPion (Index)
        End If
        
' Si on ne "mange" pas ses propres pions ou si on clique sur une case vide
        If vOldCouleur <> tInfoPion(0) Or tEchiquier(Index) = "" Then
            If vJoue = False Then
                vJoue = True
                LblJoue.Caption = "Noir Joue"
            Else
                vJoue = False
                LblJoue.Caption = "Blanc Joue"
            End If
            If Deplacement(vSel(1), Index) = 1 Then
                vSel(2) = CouleurCase(vSel(1))
                vNewCouleur = CouleurCase(Index)
                tEchiquier(Index) = vSel(0)
                Image1(Index).Picture = LoadPicture("./" & vSel(0) & "Sur" & vNewCouleur & ".bmp")
                tEchiquier(vSel(1)) = ""
                Image1(vSel(1)).Picture = LoadPicture("./" & "Case" & vSel(2) & ".bmp")
            Else
                MsgBox "Ce déplacement n'est pas permis !", vbCritical, "Erreur"
                
                If vJoue = False Then
                    vJoue = True
                    LblJoue.Caption = "Noir Joue"
                Else
                    vJoue = False
                    LblJoue.Caption = "Blanc Joue"
                End If
            
            End If
            
            For vCount = 0 To 2
                vSel(vCount) = ""
            Next
        
' Sélectionne la nouvelle pièce si on clique 2X sur sa couleur
        Else
            vSel(1) = Index
            vSel(0) = tEchiquier(Index)
            ShpSel.Left = Image1(Index).Left
            ShpSel.Top = Image1(Index).Top
            ShpSel.Visible = True
        End If
        
    End If
    
End Sub

Private Sub MenuFichierNew_Click()
    Form_Load
End Sub

Private Sub MenuFichierQuitter_Click()
    End
End Sub

Function CouleurCase(ByVal vNum As Integer) As String
Dim vCntCoul As Integer
Dim vFlipCoul As Integer

    For vCntCoul = 0 To 63
        If vCntCoul Mod 8 = 0 Then
            If vFlipCoul = 0 Then
                vFlipCoul = 1
            Else
                vFlipCoul = 0
            End If
        End If
        If vCntCoul = vNum Then
            If (vCntCoul + vFlipCoul) Mod 2 = 1 Then
                CouleurCase = "Blanc"
            Else
                CouleurCase = "Noir"
            End If
        End If
    Next
End Function

Function InfoPion(ByVal vNum As Integer) As String
    If Right(tEchiquier(vNum), 1) = "r" Then
        tInfoPion(0) = "Noir"
        tInfoPion(1) = Left(tEchiquier(vNum), Len(tEchiquier(vNum)) - 4)
    Else
        tInfoPion(0) = "Blanc"
        tInfoPion(1) = Left(tEchiquier(vNum), Len(tEchiquier(vNum)) - 5)
    End If
End Function

Function Deplacement(ByVal vStart As Integer, vStop As Integer) As Integer
Dim tCoordonnes(1, 1) As Integer
Dim vCountPion As Integer
    
    tCoordonnes(0, 0) = vStop Mod 8                         ' Colone Stop
    tCoordonnes(0, 1) = (vStop - tCoordonnes(0, 0)) / 8     ' Ligne Stop
    tCoordonnes(1, 0) = vStart Mod 8                        ' Colone Start
    tCoordonnes(1, 1) = (vStart - tCoordonnes(1, 0)) / 8    ' Ligne Start

    InfoPion (vStart)
    
    If tInfoPion(1) = "Pion" Then
        If tInfoPion(0) = "Blanc" Then
            If vStart >= 48 And vStart <= 55 Then
                If vStart - vStop = 16 Or vStart - vStop = 8 Then
                    Deplacement = 1
                End If
            Else
                If vStart - vStop = 8 Then
                    Deplacement = 1
                End If
            End If
        Else
            If vStart >= 8 And vStart <= 15 Then
                If vStop - vStart = 16 Or vStop - vStart = 8 Then
                    Deplacement = 1
                End If
            Else
                If vStop - vStart = 8 Then
                    Deplacement = 1
                End If
            End If
        End If
            
    ElseIf tInfoPion(1) = "Tour" Then
        If (vStop - vStart) Mod 8 = 0 Or tCoordonnes(1, 1) = tCoordonnes(0, 1) Then
            Deplacement = 1
            For vCountPion = tCoordonnes(1, 1) + 1 To tCoordonnes(0, 1)
                If tEchiquier(vCountPion * 8 + tCoordonnes(0, 0)) <> "" Then
                    Deplacement = 0
                End If
            Next
            
        End If
            
    ElseIf tInfoPion(1) = "Cheval" Then
        If (Abs(tCoordonnes(1, 0) - tCoordonnes(0, 0)) = 1 And Abs(tCoordonnes(1, 1) - tCoordonnes(0, 1)) = 2) Or (Abs(tCoordonnes(1, 0) - tCoordonnes(0, 0)) = 2 And Abs(tCoordonnes(1, 1) - tCoordonnes(0, 1)) = 1) Then
            Deplacement = 1
        End If
            
    ElseIf tInfoPion(1) = "Fou" Then
        If Abs(vStop - vStart) Mod 9 = 0 Or Abs(vStop - vStart) Mod 7 = 0 Then
            Deplacement = 1
        End If
            
    ElseIf tInfoPion(1) = "Reine" Then
        If Abs(vStop - vStart) Mod 9 = 0 Or Abs(vStop - vStart) Mod 7 = 0 Or Abs(vStop - vStart) Mod 8 = 0 Or tCoordonnes(1, 1) = tCoordonnes(0, 1) Then
            Deplacement = 1
        End If
            
    ElseIf tInfoPion(1) = "Roi" Then
        If Abs(vStart - vStop) = 1 Or Abs(vStart - vStop) = 7 Or Abs(vStart - vStop) = 8 Or Abs(vStart - vStop) = 9 Then
            Deplacement = 1
        End If
            
    End If
End Function
