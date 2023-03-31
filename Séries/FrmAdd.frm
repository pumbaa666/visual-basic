VERSION 5.00
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajouter"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNever 
      Caption         =   "&Jamais"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton CmdAll 
      Caption         =   "Toute la &semaine"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton CmdOuvrable 
      Caption         =   "Tout les jours &ouvrables"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "dimanche"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "samedi"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "vendredi"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "jeudi"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "mercredi"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "mardi"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox ChkJours 
      Caption         =   "lundi"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton CmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox TxtHeure 
      Height          =   285
      Left            =   960
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "hh:mm:ss"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox TxtTitre 
      Height          =   285
      Left            =   960
      MaxLength       =   80
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Jours :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Heure :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Titre :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAll_Click()
Dim vCount As Integer
    For vCount = 0 To 6
        ChkJours(vCount).Value = 1
    Next
End Sub

Private Sub CmdAnnuler_Click()
Dim vCount As Integer
    TxtTitre.Text = ""
    TxtHeure.Text = "hh:mm:ss"
    For vCount = 0 To 6
        ChkJours(vCount).Value = 0
    Next
    FrmMain.Show
    FrmAdd.Hide
End Sub

Private Sub CmdNever_Click()
Dim vCount As Integer
    For vCount = 0 To 6
        ChkJours(vCount).Value = 0
    Next
End Sub

Private Sub CmdOk_Click()
Dim sStruct As Structure
Dim vCount As Integer
    If TxtTitre.Text = "" Then
        MsgBox "Entrer le titre!", vbCritical, "Erreur"
    ElseIf TxtHeure.Text = "hh:mm:ss" Then
        MsgBox "Entrer l'heure!", vbCritical, "Erreur"
    ElseIf Int(Left(TxtHeure.Text, 2)) > 23 Or Int(Mid(TxtHeure.Text, 4, 2)) > 59 Or Int(Right(TxtHeure.Text, 2)) > 59 Or Len(TxtHeure.Text) <> 8 Then
        MsgBox "le format de l'heure n'est pas correct!", vbCritical, "Erreur"
    ElseIf ChkJours(0).Value <> Checked And ChkJours(1).Value <> Checked And ChkJours(2).Value <> Checked And ChkJours(3).Value <> Checked And ChkJours(4).Value <> Checked And ChkJours(5).Value <> Checked And ChkJours(6).Value <> Checked Then
        MsgBox "Choisissez les jours durant lesquels passent la série!", vbCritical, "Erreur"
    Else
        sStruct.vTitre = TxtTitre.Text
        sStruct.vHeure = TxtHeure.Text
        For vCount = 0 To 6
            sStruct.vJours(vCount) = ChkJours(vCount).Value
        Next
        
        Open "c:\temp\series.dat" For Random As #1 Len = Len(sStruct)
        Put #1, vNbEnreg, sStruct
        Close #1
        
        vNbEnreg = vNbEnreg + 1
        
        Open "c:\temp\nbseries.dat" For Output As #1
        Print #1, vNbEnreg
        Close #1
        
        CmdAnnuler_Click
    End If
End Sub

Private Sub CmdOuvrable_Click()
Dim vCount As Integer
    For vCount = 0 To 4
        ChkJours(vCount).Value = 1
    Next
    ChkJours(5).Value = 0
    ChkJours(6).Value = 0
End Sub

Private Sub TxtHeure_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If KeyAscii < 48 Or KeyAscii > 58 Then
            MsgBox "Veuillez n'entrer que des chiffres!", vbCritical, "Erreur"
            KeyAscii = 0
        End If
    End If
End Sub
