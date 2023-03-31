VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion des contacts"
   ClientHeight    =   5190
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dial1 
      Left            =   4320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Open File Contact"
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   7455
      Begin VB.Label LblFichier 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contact"
      Height          =   4095
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   3600
         TabIndex        =   33
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdCreat 
         Caption         =   "C&reate"
         Height          =   375
         Left            =   3600
         TabIndex        =   31
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "Ne&xt"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   "&Preview"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox TxtMail 
         Height          =   285
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   17
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox TxtPhone 
         Height          =   285
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   16
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox TxtCountry 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox TxtCity 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   14
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox TxtNPA 
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox TxtAdress 
         Height          =   285
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   12
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox TxtName 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox TxtFirst 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxtNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "(*) Must be fill"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Mail :"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone :"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Country :"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "City :"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NPA :"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Address :"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "(*) Name :"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "(*) First name :"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "N° d'enregistrement :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actions"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1815
      Begin VB.CommandButton CmdQuit 
         Caption         =   "&Quit"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton CmdTri 
         Caption         =   "Order &A to Z"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton CmdModify 
         Caption         =   "&Modify"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New Character"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton CmdFind 
         Caption         =   "Fin&d"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuFileNewFile 
         Caption         =   "New Fi&le"
      End
      Begin VB.Menu MenuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileNewCharacter 
         Caption         =   "&New  Character"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFileFind 
         Caption         =   "Fin&d"
         Shortcut        =   ^F
      End
      Begin VB.Menu MenuFileDelete 
         Caption         =   "&Delete"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu MenuFileModify 
         Caption         =   "&Modify"
         Shortcut        =   ^M
      End
      Begin VB.Menu MenuFileTri 
         Caption         =   "Order &A to Z"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuOption 
      Caption         =   "&Option"
      Begin VB.Menu MenuOptionAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCntOpen As Integer
Dim vTest As Integer

Private Sub CmdCancel_Click()
Dim sCancel As Struct_Contact
    If vTest <> 6 Then
        Open LblFichier.Caption For Random As #1 Len = Len(sCancel)
        Get #1, TxtNo.Text - 1, sCancel
        Close #1
        TxtNo.Text = sCancel.ID
        TxtFirst.Text = Trim(sCancel.FirstName)
        TxtName.Text = Trim(sCancel.Name)
        TxtAdress.Text = Trim(sCancel.Adress)
        TxtNPA.Text = Trim(sCancel.NPA)
        TxtCity.Text = Trim(sCancel.City)
        TxtCountry.Text = Trim(sCancel.Country)
        TxtPhone.Text = Trim(sCancel.Phone)
        TxtMail.Text = Trim(sCancel.Mail)
    Else
        CmdPrev_Click
    End If
            
    CmdPrev.Visible = True
    CmdNext.Visible = True
    CmdCancel.Visible = False
    CmdCreat.Visible = False
    SHTextBox ("False")
    SHCmd ("True")
End Sub

Private Sub CmdCreat_Click()
Dim sSave As Struct_Contact
    If TxtFirst.Text = "" Or TxtName.Text = "" Then
        MsgBox "Il manque des paramètres!", vbCritical, "Erreur"
    Else
        sSave.ID = TxtNo.Text
        sSave.FirstName = TxtFirst.Text
        sSave.Name = TxtName.Text
        sSave.Adress = TxtAdress.Text
        sSave.NPA = TxtNPA.Text
        sSave.City = TxtCity.Text
        sSave.Country = TxtCountry.Text
        sSave.Phone = TxtPhone.Text
        sSave.Mail = TxtMail.Text
        Open Dial1.FileName For Random As #1 Len = Len(sSave)
        Put #1, vCntOpen, sSave
        Close #1
        CmdCreat.Visible = False
        CmdCancel.Visible = False
        CmdNext.Visible = True
        CmdPrev.Visible = True
        
        SHTextBox ("False")
        SHCmd ("True")
    End If
End Sub

Private Sub CmdDelete_Click()
Dim vYes As Integer
Dim sDel As Struct_Contact
Dim vCount As Integer
Dim vTemp As String
    vYes = MsgBox("Etes-vous sur de vouloir supprimer cette personne?!?", vbYesNo, "Suppression")
    If vYes = vbYes Then
        Open LblFichier.Caption For Random As #1 Len = Len(sDel)
        Open "c:\temp.tmp" For Random As #2 Len = Len(sDel)
        For vCount = 1 To LOF(1) / Len(sDel)
            If vCount < Int(TxtNo.Text) Then
                Get #1, vCount, sDel
                Put #2, vCount, sDel
            ElseIf vCount <> Int(TxtNo.Text) Then
                Get #1, vCount, sDel
                sDel.ID = vCount - 1
                Put #2, vCount - 1, sDel
            End If
        Next
        
        Close #1
        Close #2
        Kill (LblFichier.Caption)
        FileCopy "c:\temp.tmp", LblFichier.Caption
        Kill ("c:\temp.tmp")
        
        If Int(TxtNo.Text) = 1 Then
            CmdNext_Click
            CmdPrev_Click
            CmdNext_Click
            CmdPrev_Click
        Else
            CmdPrev_Click
        End If
        
        Open LblFichier.Caption For Random As #1 Len = Len(sDel)
        If LOF(1) = 0 Then
            ClearBox
            TxtNo.Text = ""
            SHTextBox ("False")
            SHCmd ("False")
        ElseIf LOF(1) / Len(sDel) = 1 Then
            TxtNo.Text = "1"
        End If
        Close #1
    End If
End Sub

Private Sub CmdFind_Click()
    ClearBox
    TxtNo.Text = ""
    SHTextBox ("False")
    SHCmd ("False")
    TxtName.Enabled = True
    TxtFirst.Enabled = True
    TxtFirst.SetFocus
    CmdNext.Visible = False
    CmdPrev.Visible = False
    CmdSearch.Visible = True
    CmdCancel.Visible = True
End Sub

Private Sub CmdModify_Click()
    vTest = 6
    SHTextBox ("True")
    SHCmd ("False")
    CmdNext.Visible = False
    CmdNext.Visible = False
    CmdCancel.Visible = True
    CmdCreat.Visible = True
End Sub

Private Sub CmdNew_Click()
    Do
        CmdNext_Click
    Loop While (vTest <> 2)
    
    vCntOpen = vCntOpen + 1
    ClearBox
    SHTextBox ("True")
    SHCmd ("False")
    CmdPrev.Visible = False
    CmdNext.Visible = False
    CmdCancel.Visible = True
    CmdCreat.Visible = True
End Sub

Private Sub CmdNext_Click()
    vCntOpen = vCntOpen + 1
    vTest = 1
    MenuFileOpen_Click
End Sub

Private Sub CmdPrev_Click()
    If vCntOpen <> 1 Then
        vCntOpen = vCntOpen - 1
        vTest = 1
        MenuFileOpen_Click
    End If
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub CmdSearch_Click()
Dim sSearch As Struct_Contact
Dim vCount As Integer
Dim tLen(1) As Integer
    If TxtFirst.Text = "" And TxtName.Text = "" Then
        MsgBox "Veuillez entrer le nom ou le prénom", vbCritical, "Erreur"
    Else
        tLen(0) = Len(TxtFirst.Text)
        tLen(1) = Len(TxtName.Text)
        Open LblFichier.Caption For Random As #1 Len = Len(sSearch)
        For vCount = 1 To LOF(1) / Len(sSearch)
            Get #1, vCount, sSearch
            If TxtFirst.Text <> "" And TxtName.Text <> "" Then
                If LCase(Trim(Left(sSearch.FirstName, tLen(0)))) = LCase(Trim(TxtFirst)) Or LCase(Trim(Left(sSearch.Name, tLen(1)))) = LCase(Trim(TxtName.Text)) Then
                    TxtNo.Text = Trim(sSearch.ID)
                    TxtFirst.Text = Trim(sSearch.FirstName)
                    TxtName.Text = Trim(sSearch.Name)
                    TxtAdress.Text = Trim(sSearch.Adress)
                    TxtNPA.Text = Trim(sSearch.NPA)
                    TxtCity.Text = Trim(sSearch.City)
                    TxtCountry.Text = Trim(sSearch.Country)
                    TxtPhone.Text = Trim(sSearch.Phone)
                    TxtMail.Text = Trim(sSearch.Mail)
                    vCntOpen = sSearch.ID
                    Exit For
                End If
            Else
                If LCase(Trim(Left(sSearch.FirstName, tLen(0)))) = LCase(Trim(TxtFirst)) And LCase(Trim(Left(sSearch.Name, tLen(1)))) = LCase(Trim(TxtName.Text)) Then
                    TxtNo.Text = Trim(sSearch.ID)
                    TxtFirst.Text = Trim(sSearch.FirstName)
                    TxtName.Text = Trim(sSearch.Name)
                    TxtAdress.Text = Trim(sSearch.Adress)
                    TxtNPA.Text = Trim(sSearch.NPA)
                    TxtCity.Text = Trim(sSearch.City)
                    TxtCountry.Text = Trim(sSearch.Country)
                    TxtPhone.Text = Trim(sSearch.Phone)
                    TxtMail.Text = Trim(sSearch.Mail)
                    vCntOpen = sSearch.ID
                    Exit For
                End If
            End If
        Next
        Close #1
        SHTextBox ("False")
        SHCmd ("True")
        CmdSearch.Visible = False
        CmdCancel.Visible = False
        CmdNext.Visible = True
        CmdPrev.Visible = True
        If TxtNo.Text = "" Then
            vTest = 1
            MenuFileOpen_Click
            MsgBox "Aucun résultats trouvés", vbInformation, "Recherche"
        End If
    End If
End Sub

Private Sub Form_Load()
    vCntOpen = 1
    SHTextBox ("False")
    SHCmd ("False")
End Sub

Private Sub MenuFileDelete_Click()
    CmdDelete_Click
End Sub

Private Sub MenuFileFind_Click()
    CmdFind_Click
End Sub

Private Sub MenuFileModify_Click()
    CmdModify_Click
End Sub

Private Sub MenuFileNewCharacter_Click()
    CmdNew_Click
End Sub

Private Sub MenuFileNewFile_Click()
    FrmExplorateur.Show
    FrmMain.Hide
End Sub

Private Sub MenuFileTri_Click()
    CmdTri_Click = True
End Sub

Private Sub MenuFileOpen_Click()
Dim sSave As Struct_Contact
Dim vLngFile As Integer
    CmdNext.Enabled = True
    CmdPrev.Enabled = True
    If vTest <> 1 Then
        Dial1.ShowOpen
    Else
        vTest = 0
    End If
    
    If Dial1.FileName <> "" Then
        Close #1
        Open Dial1.FileName For Random As #1 Len = Len(sSave)
        Get #1, vCntOpen, sSave
        vLngFile = Len(sSave)
        Close #1
        LblFichier.Caption = Dial1.FileName
        
        If sSave.ID = 0 Then
            If vLngFile <> 0 Then
                vCntOpen = vCntOpen - 1
                vTest = 2
            Else
                CmdNew.Enabled = True
            End If
        Else
            TxtNo.Text = sSave.ID
            TxtFirst.Text = Trim(sSave.FirstName)
            TxtName.Text = Trim(sSave.Name)
            TxtAdress.Text = Trim(sSave.Adress)
            TxtNPA.Text = Trim(sSave.NPA)
            TxtCity.Text = Trim(sSave.City)
            TxtCountry.Text = Trim(sSave.Country)
            TxtPhone.Text = Trim(sSave.Phone)
            TxtMail.Text = Trim(sSave.Mail)
            SHTextBox ("False")
            SHCmd ("True")
        End If
    End If
End Sub

Private Sub MenuFileQuit_Click()
    End
End Sub

Private Sub MenuFileSave_Click()
Dim sSave As Struct_Contact
    If TxtFirst.Text <> "" And TxtName.Text <> "" Then
        Dial1.ShowSave
        vCntOpen = vCntOpen + 1
        TxtNo.Text = vCntOpen
        
        sSave.ID = TxtNo.Text
        sSave.FirstName = TxtFirst.Text
        sSave.Name = TxtName.Text
        sSave.Adress = TxtAdress.Text
        sSave.NPA = TxtNPA.Text
        sSave.City = TxtCity.Text
        sSave.Country = TxtCountry.Text
        sSave.Phone = TxtPhone.Text
        sSave.Mail = TxtMail.Text
        
        Open Dial1.FileName For Random As #1 Len = Len(sSave)
        Put #1, vCntOpen, sSave
        Close #1
       
        LblFichier.Caption = Dial1.FileName
    Else: MsgBox "Il manque des paramètres", vbCritical, "Erreur"
    End If
End Sub

Function ClearBox()
    TxtNo.Text = vCntOpen
    TxtFirst = ""
    TxtName = ""
    TxtAdress = ""
    TxtNPA = ""
    TxtCity = ""
    TxtCountry = ""
    TxtPhone = ""
    TxtMail = ""
End Function

Private Sub MenuOptionAbout_Click()
    FrmAbout.Show
    FrmMain.Hide
End Sub

Function SHTextBox(ByVal TF As String)
    TxtFirst.Enabled = TF
    TxtName.Enabled = TF
    TxtAdress.Enabled = TF
    TxtNPA.Enabled = TF
    TxtCity.Enabled = TF
    TxtCountry.Enabled = TF
    TxtPhone.Enabled = TF
    TxtMail.Enabled = TF
End Function

Function SHCmd(ByVal TF As String)
    CmdFind.Enabled = TF
    CmdNew.Enabled = TF
    CmdDelete.Enabled = TF
    CmdModify.Enabled = TF
    CmdTri.Enabled = TF
    
    MenuFileFind.Enabled = TF
    MenuFileNewCharacter.Enabled = TF
    MenuFileDelete.Enabled = TF
    MenuFileModify.Enabled = TF
    MenuFileTri.Enabled = TF
    MenuFileSave.Enabled = TF
End Function

'Function StructToText(ByVal sFonction As Struct_Contact)
'    TxtNo.Text = sFonction.ID
'    TxtFirst.Text = Trim(sFonction.FirstName)
'    TxtName.Text = Trim(sFonction.Name)
'    TxtAdress.Text = Trim(sFonction.Adress)
'    TxtNPA.Text = Trim(sFonction.NPA)
'    TxtCity.Text = Trim(sFonction.City)
'    TxtCountry.Text = Trim(sFonction.Country)
'    TxtPhone.Text = Trim(sFonction.Phone)
'    TxtMail.Text = Trim(sFonction.Mail)
'End Function

Private Sub TxtNPA_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
        MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur"
        KeyAscii = 0
    End If
End Sub
