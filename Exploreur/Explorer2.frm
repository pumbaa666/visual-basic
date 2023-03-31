VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Explorateur"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "EXPLOR~1.frx":0000
      Left            =   3600
      List            =   "EXPLOR~1.frx":0016
      TabIndex        =   7
      Text            =   "*.*"
      Top             =   3120
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton CmdOuvrir 
      Caption         =   "&Ouvrir"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtChemin 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label LblLecteur 
      Caption         =   "Lecteur"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label LblType 
      Caption         =   "Type de fichiers"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Chemin"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vLarg As Integer
Dim vHaut As Integer

Private Sub CmdAnnuler_Click()
    End
End Sub

Private Sub CmdOuvrir_Click()
Dim vOui As Integer
Dim vProg As String
' LCase converti la chaine de caractère en minuscule et
' Right regarde le nb de caractères demandé depuis la droite
' Shell ouvre le programme choisi
    If LCase(Right(TxtChemin.Text, 3)) = "doc" Then
        vProg = "C:\Program Files\Microsoft Office\Office\WINWORD.EXE "
    ElseIf LCase(Right(TxtChemin.Text, 3)) = "bmp" Then
        vProg = "C:\PROGRA~1\ACCESS~1\MSPAINT.EXE "
    ElseIf LCase(Right(TxtChemin.Text, 3)) = "xls" Then
        vProg = "C:\Program Files\Microsoft Office\Office\EXCEL.EXE "
    ElseIf LCase(Right(TxtChemin.Text, 3)) = "txt" Or LCase(Right(TxtChemin.Text, 3)) = "bat" Then
        vProg = "C:\WINDOWS\NOTEPAD.EXE "
'    ElseIf LCase(Right(TxtChemin.Text, 3)) = "vbp" Or LCase(Right(TxtChemin.Text, 3)) = "frm" Then
'        vProg = "c:\Program Files\DevStudio\VB\VB5.EXE "
    Else
        vOui = MsgBox("Le format souhaité n'est pas pris en charge par cet explorateur, voulez-vous l'ouvrir avec Notepad?!?", vbYesNo)
        If vOui = vbYes Then
            vProg = "C:\WINDOWS\NOTEPAD.EXE "
        End If
    End If
' Regarde si on est à la racine
    If Len(Dir1.Path) = 3 Then
        Var = Shell(vProg & Dir1.Path & File1.filename, vbMaximizedFocus)
    Else
        Var = Shell(vProg & Dir1.Path & "\" & File1.filename, vbMaximizedFocus)
    End If
End Sub

Private Sub combo1_Click()
    File1.Pattern = Combo1.Text
End Sub

Private Sub Dir1_Change()
    TxtChemin.Text = Dir1.Path
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    CmdOuvrir_Click
End Sub

Private Sub File1_Click()
    If Len(Dir1.Path) = 3 Then
        TxtChemin.Text = TxtChemin.Text & File1.filename
    Else
        TxtChemin.Text = Dir1.Path & "\" & File1.filename
    End If
End Sub

Private Sub Form_Load()
    Dir1.Path = "c:\"
    Drive1.Drive = "c:\"
    vLarg = Form1.Width
    vHaut = Form1.Height
End Sub

Private Sub Form_Resize()
    If Form1.Height > 2000 And Form1.Width > 3500 Then
        CmdOuvrir.Left = CmdOuvrir.Left + (Form1.Width - vLarg)
        CmdAnnuler.Left = CmdAnnuler.Left + (Form1.Width - vLarg)
        Combo1.Left = Combo1.Left + (Form1.Width - vLarg)
        Drive1.Left = Drive1.Left + (Form1.Width - vLarg)
        TxtChemin.Width = TxtChemin.Width + (Form1.Width - vLarg)
        File1.Left = File1.Left + (Form1.Width - vLarg) / 2
        File1.Width = File1.Width + (Form1.Width - vLarg) / 2
        Dir1.Width = Dir1.Width + (Form1.Width - vLarg) / 2
        LblType.Left = LblType.Left + (Form1.Width - vLarg)
        LblLecteur.Left = LblLecteur.Left + (Form1.Width - vLarg)
    
        CmdOuvrir.Top = CmdOuvrir.Top + (Form1.Height - vHaut)
        CmdAnnuler.Top = CmdAnnuler.Top + (Form1.Height - vHaut)
        Combo1.Top = Combo1.Top + (Form1.Height - vHaut)
        Drive1.Top = Drive1.Top + (Form1.Height - vHaut)
        LblType.Top = LblType.Top + (Form1.Height - vHaut)
        LblLecteur.Top = LblLecteur.Top + (Form1.Height - vHaut)
        Dir1.Height = Dir1.Height + (Form1.Height - vHaut)
        File1.Height = File1.Height + (Form1.Height - vHaut)
  
        vLarg = Form1.Width
        vHaut = Form1.Height
    Else: MsgBox "Vous ne pouvez pas réduire autant la fenêtre", vbCritical, "Errorssssss!!!"
        Form1.Height = vHaut
        Form1.Width = vLarg
    End If
End Sub

Private Sub TxtChemin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
' Permet de rentrer manuellement le chemin
        Dir1.Path = TxtChemin.Text
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
' Permet de rentrer manuellementle type de fichier a ouvrir
         File1.Pattern = Combo1.Text
    End If
End Sub
