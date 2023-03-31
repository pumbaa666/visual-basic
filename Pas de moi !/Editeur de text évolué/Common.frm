VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Common 
   BorderStyle     =   0  'None
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.TreeView TreeView 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6376
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PN3.Button ButtonOk 
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   6050618
      BackColor       =   15526369
      Caption         =   "Ouvir"
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3255
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList"
      ColHdrIcons     =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fichier"
         Object.Width           =   4410
         ImageIndex      =   10
      EndProperty
   End
   Begin PN3.Button ButtonCancel 
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Appearance      =   0
      ColorFlat       =   192
      BackColor       =   14938367
      Caption         =   "Annuler"
      ForeColor       =   192
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":1D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":2B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":4866
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":6570
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":73C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":8214
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":9F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":BC28
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":CA7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":D354
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":DC2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Common.frx":E508
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelTitle 
      BackStyle       =   0  'Transparent
      Caption         =   " Ouvrir un fichier texte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   7335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00B4A587&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00D8D0C0&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Fso As New FileSystemObject
Dim DriveZ As Drive
Dim FolderZ As Folder
Dim FileZ As File
Dim I As Integer
Dim Node As Node
Dim Dossier, Tampon
Dim IsCancel As Boolean
Dim IsOk As Boolean

Private Sub ButtonCancel_Click()

IsCancel = True

End Sub

Private Sub ButtonOk_Click()

IsOk = True

End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Pour bouger la fenêtre
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub Form_Load()

ListView.ColumnHeaders.Item(1).Width = ListView.Width

'Recherche des périphériques

TreeView.Nodes.Add , , "Root", "Poste de travail", 1

For Each DriveZ In Fso.Drives
    Select Case DriveZ.DriveType
        Case Removable
            TreeView.Nodes.Add "Root", 4, DriveZ.DriveLetter & ":\", "Disquette (" & DriveZ.DriveLetter & ":)", 2
            If Fso.FolderExists(DriveZ.DriveLetter & ":\") = True Then TreeView.Nodes.Add DriveZ.DriveLetter & ":\", 4, DriveZ.DriveLetter & ":\" & "Temp\"
        Case Fixed
            TreeView.Nodes.Add "Root", 4, DriveZ.DriveLetter & ":\", DriveZ.VolumeName & " (" & DriveZ.DriveLetter & ":)", 3
            If Fso.FolderExists(DriveZ.DriveLetter & ":\") = True Then TreeView.Nodes.Add DriveZ.DriveLetter & ":\", 4, DriveZ.DriveLetter & ":\" & "Temp\"
        Case Remote
            TreeView.Nodes.Add "Root", 4, DriveZ.DriveLetter & ":\", DriveZ.VolumeName & " (" & DriveZ.DriveLetter & ":)", 5
            If Fso.FolderExists(DriveZ.DriveLetter & ":\") = True Then TreeView.Nodes.Add DriveZ.DriveLetter & ":\", 4, DriveZ.DriveLetter & ":\" & "Temp\"
        Case RamDisk
            TreeView.Nodes.Add "Root", 4, DriveZ.DriveLetter & ":\", DriveZ.VolumeName & " (" & DriveZ.DriveLetter & ":)", 6
            If Fso.FolderExists(DriveZ.DriveLetter & ":\") = True Then TreeView.Nodes.Add DriveZ.DriveLetter & ":\", 4, DriveZ.DriveLetter & ":\" & "Temp\"
        Case CDRom
            TreeView.Nodes.Add "Root", 4, DriveZ.DriveLetter & ":\", "Lecteur CD (" & DriveZ.DriveLetter & ":)", 4
            If Fso.FolderExists(DriveZ.DriveLetter & ":\") = True Then TreeView.Nodes.Add DriveZ.DriveLetter & ":\", 4, DriveZ.DriveLetter & ":\" & "Temp\"
        Case Unknown
            TreeView.Nodes.Add "Root", 4, DriveZ.DriveLetter & ":\", "Non Reconnu (" & DriveZ.DriveLetter & ":)", 9
            If Fso.FolderExists(DriveZ.DriveLetter & ":\") = True Then TreeView.Nodes.Add DriveZ.DriveLetter & ":\", 4, DriveZ.DriveLetter & ":\" & "Temp\"
    End Select
Next DriveZ

TreeView.Nodes(1).Expanded = True

End Sub

Private Sub ListView_DblClick()

IsOk = True

End Sub

Private Sub TreeView_Click()

ListView.ListItems.Clear

If TreeView.SelectedItem.Key <> "Root" Then
    For Each FileZ In Fso.GetFolder(TreeView.SelectedItem.Key).Files
        Select Case Right(FileZ.Name, 3)
            Case "txt", "log", "rtx", "wtx"
                ListView.ListItems.Add 1, , FileZ.Name, , 13
            Case "rtf"
                ListView.ListItems.Add 1, , FileZ.Name, , 11
            Case "doc"
                ListView.ListItems.Add 1, , FileZ.Name, , 12
        End Select
    Next FileZ
End If

End Sub

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)

If Node.Key <> "Root" And TreeView.Nodes.Item(Node.Index + 1).Key = Node.Key & "Temp\" Then
    TreeView.Nodes.Remove Node.Index + 1
    For Each FolderZ In Fso.GetFolder(Node.Key).SubFolders
        TreeView.Nodes.Add Node.Key, 4, FolderZ & "\", FolderZ.Name, 7, 8
        If Fso.FolderExists(Dossier & "\") = True Then TreeView.Nodes.Add FolderZ & "\", 4, FolderZ & "\" & "Temp" & "\", "Temp"
    Next FolderZ
End If

End Sub

Public Function ShowCommon(Titre As String, Optional Répertoire As String) As String

IsOk = False
IsCancel = False

Main.Enabled = False
LabelTitle.Caption = " " & Titre
Me.Show

'On tourne en attendant l'appui sur un bouton
Do While IsOk = False And IsCancel = False
    DoEvents
Loop

'On renvoi la réponse
If IsOk = True Then ShowCommon = TreeView.SelectedItem.Key & ListView.SelectedItem.Text

'On redonne la main aux feuilles de base et on cache le Common
Me.Hide
Main.Enabled = True

End Function
