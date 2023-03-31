VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loïc Media Player"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListPref 
      Height          =   3765
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label LblListe 
      Alignment       =   2  'Center
      Caption         =   "Liste de mes préférences"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin MediaPlayerCtl.MediaPlayer LMP 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4530
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -950
      WindowlessVideo =   0   'False
   End
   Begin VB.Label LblTitre 
      Alignment       =   2  'Center
      Caption         =   "Loïc Media Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Menu MenuFichier 
      Caption         =   "Fichier"
      Begin VB.Menu MenuFichierOuvrir 
         Caption         =   "Ouvrir"
      End
      Begin VB.Menu MenuFichierTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFichierShow 
         Caption         =   "Afficher mes préférences"
      End
      Begin VB.Menu MenuFichierHide 
         Caption         =   "Masquer mes préférences"
      End
      Begin VB.Menu MenuFichierTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFichierQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu MenuLecture 
      Caption         =   "Lecture"
      Begin VB.Menu MenuLectureAleatoire 
         Caption         =   "Aléatoire"
      End
      Begin VB.Menu MenuLectureRepeter 
         Caption         =   "Répéter"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sLoad As StructListe
Dim vCount As Integer
    On Error Resume Next
    Open "c:\temp\prefmedia.dat" For Random As #1 Len = Len(sLoad)
    vNbPref = FileLen("c:\temp\prefmedia.dat") / Len(sLoad)
    For vCount = 1 To FileLen("c:\temp\prefmedia.dat") / Len(sLoad)
        Get #1, vCount, sLoad
        If sLoad.vTitre <> "" Then
            ListPref.AddItem Trim(sLoad.vTitre)
            tMusique(0, vCount - 1) = Trim(sLoad.vPath)
            tMusique(1, vCount - 1) = Trim(sLoad.vTitre)
        End If
    Next
    Close #1
End Sub

Private Sub ListPref_DblClick()
Dim sLaunch As StructListe
    LMP.AutoSize = False
    Open "c:\temp\prefmedia.dat" For Random As #1 Len = Len(sLaunch)
    Get #1, ListPref.ListIndex + 1, sLaunch
    Close #1
    LMP.Open Trim(sLaunch.vPath) & Trim(sLaunch.vTitre)
    vNumMus = ListPref.ListIndex + 1
End Sub

Private Sub ListPref_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sDel As StructListe
Dim vCount As Integer
    If ListPref.Text <> "" And (KeyCode = 8 Or KeyCode = 46) Then
        Open "c:\temp\prefmedia.dat" For Random As #1 Len = Len(sDel)
        Open "c:\temp.tmp" For Random As #2 Len = Len(sDel)
        For vCount = 1 To vNbPref
            If vCount < Int(ListPref.ListIndex) + 1 Then
                Get #1, vCount, sDel
                Put #2, vCount, sDel
            ElseIf vCount <> Int(ListPref.ListIndex) + 1 Then
                Get #1, vCount, sDel
                Put #2, vCount - 1, sDel
            End If
        Next
        
        Close #1
        Close #2
        Kill ("c:\temp\prefmedia.dat")
        FileCopy "c:\temp.tmp", "c:\temp\prefmedia.dat"
        Kill ("c:\temp.tmp")
        
        tMusique(0, ListPref.ListIndex - 1) = ""
        tMusique(1, ListPref.ListIndex - 1) = ""
        ListPref.RemoveItem (ListPref.ListIndex)
        vNbPref = vNbPref - 1
    End If
End Sub


Private Sub LMP_EndOfStream(ByVal Result As Long)
'    LMP.Volume = 100
    If MenuLectureAleatoire.Checked = False Then
        If vNumMus = ListPref.ListCount And MenuLectureRepeter.Checked = True Then
            vNumMus = 1
            LMP.Open tMusique(0, 0) & tMusique(1, 0)
        Else
            vNumMus = vNumMus + 1
            LMP.Open tMusique(0, vNumMus - 1) & tMusique(1, vNumMus - 1)
        End If
        LMP.Width = 4530
        LMP.Height = 3735
    Else
        vNumMus = Int(Rnd * ListPref.ListCount) + 1
        LMP.Open tMusique(0, vNumMus - 1) & tMusique(1, vNumMus - 1)
    End If
End Sub

Private Sub MenuFichierHide_Click()
    ListPref.Visible = False
    FrmMain.Width = 5400
    LblListe.Visible = False
End Sub

Private Sub MenuFichierOuvrir_Click()
    FrmOuvrir.Show
    FrmMain.Hide
End Sub

Private Sub MenuFichierQuitter_Click()
    End
End Sub

Private Sub MenuFichierShow_Click()
    ListPref.Visible = True
    FrmMain.Width = 8300
    LblListe.Visible = True
End Sub

Private Sub MenuLectureAleatoire_Click()
    If MenuLectureAleatoire.Checked = True Then
        MenuLectureAleatoire.Checked = False
    Else
        MenuLectureAleatoire.Checked = True
    End If
End Sub

Private Sub MenuLectureRepeter_Click()
    If MenuLectureRepeter.Checked = True Then
        MenuLectureRepeter.Checked = False
    Else
        MenuLectureRepeter.Checked = True
    End If
End Sub
