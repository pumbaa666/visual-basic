VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio et TV sur le net..."
   ClientHeight    =   1665
   ClientLeft      =   2010
   ClientTop       =   1410
   ClientWidth     =   9795
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   9795
   Begin VB.Frame Frame1 
      Caption         =   "Ecouter vos radios sur le NET ! Regardez la TV sur le NET !"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.Frame Frame2 
         Caption         =   "Choix de la radio ou de la télévision :"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton Command6 
            Caption         =   "Arreter !"
            Height          =   315
            Left            =   7920
            TabIndex        =   13
            Top             =   550
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Form1.frx":044A
            Left            =   6000
            List            =   "Form1.frx":045D
            TabIndex        =   12
            Text            =   "Télévisions"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Regarder !"
            Height          =   315
            Left            =   7920
            TabIndex        =   11
            Top             =   180
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "A propos..."
            Height          =   315
            Left            =   3840
            TabIndex        =   10
            Top             =   180
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":04B3
            Left            =   240
            List            =   "Form1.frx":04EA
            TabIndex        =   9
            Text            =   "Radios"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Quitter"
            Height          =   315
            Left            =   3840
            TabIndex        =   7
            Top             =   550
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Arreter !"
            Height          =   315
            Left            =   2400
            TabIndex        =   4
            Top             =   550
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Ecouter !"
            Height          =   315
            Left            =   2400
            TabIndex        =   3
            Top             =   180
            Width           =   1335
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   5280
            Picture         =   "Form1.frx":05AF
            Top             =   120
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label5 
            Caption         =   "Offline"
            Height          =   255
            Left            =   5280
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   5280
            Picture         =   "Form1.frx":09F1
            Top             =   120
            Width           =   480
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   8895
         ExtentX         =   15690
         ExtentY         =   10186
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Label Label4 
         Caption         =   "www.tutoriaux-z980x.fr.st"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7680
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "z980x@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

Command5.Enabled = False
Combo2.Enabled = False
If Combo1.Text = "Radios" Then
MsgBox "Vous devez sélectionner une radio !", vbInformation, "Radios and Tv on Web by z980x"
Else
Frame1.Height = 7335
Form1.Height = 8010
WebBrowser1.Visible = True
Combo1.Enabled = False
If Combo1.Text = "Radios" Then
MsgBox "Vous devez sélectionner une radio !", vbInformation, "Radios and Tv on Web by z980x"
End If

If Combo1.Text = "France Inter" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/franceinter/finter_launch_V1.html"
WebBrowser1.Width = 9255
WebBrowser1.Height = 3600
End If
If Combo1.Text = "Nostalgie" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/nostalgie/nostv2_launch.html"

End If
If Combo1.Text = "RTL" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://yacast.rtl.fr/V4/rtl/rtl_launch_V3.html"
WebBrowser1.Width = 9255
WebBrowser1.Height = 3600
End If

If Combo1.Text = "Contact FM" Then
Naviguer = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & "http://www.contactfm.com/fsRows.php")
Frame1.Height = 1455
Form1.Height = 2130
WebBrowser1.Visible = False
End If
If Combo1.Text = "Chérie FM" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/cheriefm/cheriefm_launch_V3.html"
WebBrowser1.Width = 9255
WebBrowser1.Height = 7000
End If
If Combo1.Text = "Europe 1" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/europe1/europe1_launch_V1.html"
WebBrowser1.Width = 9255
WebBrowser1.Height = 4100
End If
If Combo1.Text = "Europe 2" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://europe2.comfm.com/europe2/player.php"
WebBrowser1.Width = 9255
WebBrowser1.Height = 4100
End If
If Combo1.Text = "France Info" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/franceinfo/finfo_launch_V1.html"
WebBrowser1.Width = 7095
WebBrowser1.Height = 3495
End If
If Combo1.Text = "France Culture" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/franceculture/fculture_launch_V1.html"
WebBrowser1.Width = 7095
WebBrowser1.Height = 3735
End If
If Combo1.Text = "France Musiques" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/francemusique/fmusique_launch_V1.html"
WebBrowser1.Width = 7095
WebBrowser1.Height = 3375
End If
If Combo1.Text = "France Bleu" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/francebleu/fbleu_launch_V1.html"
WebBrowser1.Width = 7215
WebBrowser1.Height = 4215
End If
If Combo1.Text = "Fip" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/fip/fip_launch_V1.html"
WebBrowser1.Width = 9015
WebBrowser1.Height = 3495
End If
If Combo1.Text = "Le mouv'" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/lemouv/mouv_launch_V1.html"
WebBrowser1.Width = 7095
WebBrowser1.Height = 4095
End If
If Combo1.Text = "Hector" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/hector/hector_launch_V1.html"
End If
If Combo1.Text = "Fun Radio" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/fun/fun_launch_V2.html"
WebBrowser1.Height = 3375
End If
If Combo1.Text = "Skyrock" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/skyrock/skyrock_launch_V2.html"
End If
If Combo1.Text = "Rire et Chansons" Then
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://cache.yacast.fr/V4/rireetchansons/rireetchansons_launch_V3.html"
End If
End If
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Command5.Enabled = True
Form1.Height = 2145
Frame1.Height = 1455
Combo1.Enabled = True
Combo2.Enabled = True
WebBrowser1.GoHome
WebBrowser1.Height = 10
WebBrowser1.Width = 10
WebBrowser1.Visible = False
Label5.Caption = "Offline"
Image1.Visible = True
Image2.Visible = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
MsgBox "Réalisé par z980x. Codé en Visual Basic 6 Pro. Le design, c'est pas vraiment ca, je l'admets...  Pour écouter une radio, puis une autre, il faut cliquer avant sur le bouton Arreter."
End Sub

Private Sub Command5_Click()
Command1.Enabled = False
Combo1.Enabled = False
If Combo2.Text = "Télévisions" Then
MsgBox "Vous devez sélectionner une Télévision !", vbInformation, "Radios and TV on Web by z980x"
Else
Frame1.Height = 7335
Form1.Height = 8010
WebBrowser1.Visible = True
Combo2.Enabled = False
If Combo2.Text = "Télévisions" Then
MsgBox "Vous devez sélectionner une radio !", vbInformation, "Radios and TV on Web by z980x"
End If
If Combo2.Text = "JT 8h France 2" Then
WebBrowser1.Height = 5775
WebBrowser1.Width = 8895
Frame1.Height = 7335
Form1.Height = 8010
WebBrowser1.Navigate "http://www.francetv.fr/infos/videosjt/popupjt/popupinfos8h.htm"
End If
If Combo2.Text = "JT 13h France 2" Then
WebBrowser1.Height = 5775
WebBrowser1.Width = 8895
Frame1.Height = 7335
Form1.Height = 8010
WebBrowser1.Navigate "http://www.francetv.fr/infos/videosjt/popupjt/popupinfos13h.htm"
End If
If Combo2.Text = "JT 20h France 2" Then
WebBrowser1.Height = 5775
WebBrowser1.Width = 8895
Frame1.Height = 7335
Form1.Height = 8010
WebBrowser1.Navigate "http://www.francetv.fr/infos/videosjt/popupjt/popupinfos20h.htm"
End If
If Combo2.Text = "12/14 France 3" Then
Form1.Height = 2145
Frame1.Height = 1455
WebBrowser1.Navigate "http://www.francetv.fr/regions/lienram/1214.ram"
End If
If Combo2.Text = "19/20 France 3" Then
Form1.Height = 2145
Frame1.Height = 1455
WebBrowser1.Navigate "http://www.francetv.fr/regions/lienram/1920.ram"
End If
End If
End Sub

Private Sub Command6_Click()
Command1.Enabled = True
Command5.Enabled = True
Form1.Height = 2145
Frame1.Height = 1455
Combo1.Enabled = True
Combo2.Enabled = True
WebBrowser1.GoHome
WebBrowser1.Height = 10
WebBrowser1.Width = 10
WebBrowser1.Visible = False
Label5.Caption = "Offline"
Image1.Visible = True
Image2.Visible = False
End Sub

Private Sub Label3_Click()
Envoyer = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE mailto:" & Label3.Caption, vbHide)
End Sub

Private Sub Label4_Click()
Naviguer = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & Label4.Caption)
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Image1.Visible = False
Image2.Visible = True
Label5.Caption = "On Air"
End Sub

