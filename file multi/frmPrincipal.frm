VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "ListBox"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdAGauche 
      Caption         =   "<="
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdADroite 
      Caption         =   "=>"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstDroite 
      Height          =   2595
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox lstGauche 
      Height          =   2595
      ItemData        =   "frmPrincipal.frx":0000
      Left            =   240
      List            =   "frmPrincipal.frx":0019
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblCompteDroite 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblCompteGauche 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================
'Projet : Exercice ListBox
'Auteur : Johner Olivier
'Date   : Novembre 2004
'But    : Se familiariser avec le contrôle ListBox
'       : en mode extended et multiselect.
'======================================================

Option Explicit

Private Sub cmdADroite_Click()
    If lstGauche.SelCount > 0 Then
        Dim i As Integer
        For i = lstGauche.ListCount - 1 To 0 Step -1
            If lstGauche.Selected(i) Then
                lstDroite.AddItem (lstGauche.List(i))
                lstGauche.RemoveItem (i)
            End If
        Next i
        
        ' On met a jour les deux compteurs
        lblCompteGauche.Caption = Str(lstGauche.ListCount)
        lblCompteDroite.Caption = Str(lstDroite.ListCount)
        
    End If
End Sub

Private Sub cmdAGauche_Click()
    If lstDroite.SelCount > 0 Then
        Dim i As Integer
        For i = lstDroite.ListCount - 1 To 0 Step -1
            If lstDroite.Selected(i) Then
                lstGauche.AddItem (lstDroite.List(i))
                lstDroite.RemoveItem (i)
            End If
        Next i
        
        ' On met a jour les deux compteurs
        lblCompteGauche.Caption = Str(lstGauche.ListCount)
        lblCompteDroite.Caption = Str(lstDroite.ListCount)
        
    End If
End Sub

Private Sub lstDroite_DblClick()
    lstDroite.Clear
    
    '### On met a jour les deux compteurs
    lblCompteGauche.Caption = Str(lstGauche.ListCount)
    lblCompteDroite.Caption = Str(lstDroite.ListCount)
End Sub

Private Sub lstGauche_DblClick()
    lstGauche.Clear
    
    '### On met a jour les deux compteurs
    lblCompteGauche.Caption = Str(lstGauche.ListCount)
    lblCompteDroite.Caption = Str(lstDroite.ListCount)
End Sub

Private Sub cmdAjouter_Click()
    Dim Element As String
    ' Demande a l'utilisateur d'entre un element dont on retrouve la valeur
    ' dans la variable Element
    Element = InputBox("Entrez l'élément à ajouter")
    
    ' Il n'est utile d'ajouter seulement si l'utilisateur a entre qq chose
    If Element <> "" Then
        lstGauche.AddItem (Element)
    End If

    ' On met a jour les deux compteurs
    lblCompteGauche.Caption = Str(lstGauche.ListCount)
    lblCompteDroite.Caption = Str(lstDroite.ListCount)
End Sub

Private Sub cmdSupprimer_Click()
    Dim i As Integer
    For i = lstGauche.ListCount - 1 To 0 Step -1
        If lstGauche.Selected(i) Then
            lstGauche.RemoveItem (i)
        End If
    Next i

    ' On met a jour les deux compteurs
    lblCompteGauche.Caption = Str(lstGauche.ListCount)
    lblCompteDroite.Caption = Str(lstDroite.ListCount)
End Sub

' Des le chargement de la feuille, on met a jour le nombre d'elements
' Ceci a cause du fait qu'un ListBox peut deja contenir des elements
' dans les proprietes par defauts.
Private Sub Form_Load()
    lblCompteGauche.Caption = Str(lstGauche.ListCount)
    lblCompteDroite.Caption = Str(lstDroite.ListCount)
End Sub
