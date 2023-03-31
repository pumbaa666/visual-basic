VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   180
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox TxtPhrase 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   6735
   End
   Begin VB.ComboBox ComboFin 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Text            =   "KKK"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton CmdQuitter 
      Caption         =   "&Quitter"
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox ComboKoi 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Text            =   "Koi"
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox ComboVerbe 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2280
      List            =   "Form1.frx":000A
      TabIndex        =   2
      Text            =   "Kèss"
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox ComboQui 
      Height          =   315
      ItemData        =   "Form1.frx":0016
      Left            =   360
      List            =   "Form1.frx":002C
      TabIndex        =   1
      Text            =   "Kidonk"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Générer"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOk_Click()
    TxtPhrase.Text = ComboQui.Text & " " & ComboVerbe.Text & " " & ComboKoi.Text & " " & ComboFin.Text
End Sub

Private Sub CmdQuitter_Click()
    End
End Sub

Private Sub ComboVerbe_Click()
    If ComboVerbe.ListCount = 0 Then
        ComboKoi.Clear
        ComboKoi.AddItem "le cul"
        ComboKoi.AddItem "les pieds"
        ComboKoi.AddItem "la tête"
        ComboKoi.AddItem "le petit doigt"
        ComboKoi.AddItem "les intestins"

        ComboFin.Clear
        ComboFin.AddItem "qui grattent"
        ComboFin.AddItem "enflé"
        ComboFin.AddItem "d'Homer Simpson"
        ComboFin.AddItem "agile"
        ComboFin.AddItem "plein de caca"
    Else
        ComboKoi.Clear
        ComboKoi.AddItem "un pauvre"
        ComboKoi.AddItem "tellement"
        ComboKoi.AddItem "le pire"
        ComboKoi.AddItem "purement"
        ComboKoi.AddItem "un gros"

        ComboFin.Clear
        ComboFin.AddItem "autiste"
        ComboFin.AddItem "Homer Simpson"
        ComboFin.AddItem "chemo"
        ComboFin.AddItem "con"
        ComboFin.AddItem "mouton"
    End If
End Sub

Private Sub Command1_Click()
Dim vRand As Integer
Dim vTest As Integer

    ComboQui.ListIndex = Int(Rnd * 5)
    
    vTest = Int(Rnd * 2)
    ComboVerbe.ListIndex = vTest
    If vTest = 0 Then
        ComboKoi.Clear
        ComboKoi.AddItem "le cul"
        ComboKoi.AddItem "les pieds"
        ComboKoi.AddItem "la tête"
        ComboKoi.AddItem "le petit doigt"
        ComboKoi.AddItem "les intestins"

        ComboFin.Clear
        ComboFin.AddItem "qui grattent"
        ComboFin.AddItem "enflé"
        ComboFin.AddItem "d'Homer Simpson"
        ComboFin.AddItem "agile"
        ComboFin.AddItem "plein de caca"
    Else
        ComboKoi.Clear
        ComboKoi.AddItem "un pauvre"
        ComboKoi.AddItem "tellement"
        ComboKoi.AddItem "le pire"
        ComboKoi.AddItem "purement"
        ComboKoi.AddItem "un gros"

        ComboFin.Clear
        ComboFin.AddItem "autiste"
        ComboFin.AddItem "Homer Simpson"
        ComboFin.AddItem "chemo"
        ComboFin.AddItem "con"
        ComboFin.AddItem "mouton"
    End If
    
    ComboKoi.ListIndex = Int(Rnd * 5)
    
    ComboFin.ListIndex = Int(Rnd * 5)
    
    TxtPhrase.Text = ComboQui.Text & " " & ComboVerbe.Text & " " & ComboKoi.Text & " " & ComboFin.Text
End Sub
