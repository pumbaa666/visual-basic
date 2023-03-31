VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Forme 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   240
   ClientTop       =   3345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6720
   Begin VB.Timer Timer 
      Interval        =   250
      Left            =   6120
      Top             =   120
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5106
      _Version        =   393217
      ScrollBars      =   2
      FileName        =   "C:\VB Appli\RichTextBox Test\Exemple de texte.txt"
      TextRTF         =   $"Forme.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Forme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PointAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32" ( _
                                  lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" ( _
                                  ByVal hwnd As Long, _
                                  lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                                  ByVal hwnd As Long, _
                                  ByVal wMsg As Long, _
                                  ByVal wParam As Long, _
                                  lParam As Any) As Long
Private Const EM_CHARFROMPOS As Long = &HD7
'

Private Sub Form_Resize()

    ' Positionne le RTB pour qu'il prenne toute la place
    RTB.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub RTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Moi As RECT, Souris As PointAPI, Caract As PointAPI
    Dim PosCar As Long, Temp As String
    Static Memo As PointAPI ' variable qui garde ses valeurs entre deux cycles
    
    ' Pas la peine de retester, la souris n'a pas boug�
    If Memo.x = x And Memo.y = y Then Exit Sub
    ' M�morise pour prochain tour
    Memo.x = x
    Memo.y = y
    
    ' Toutes ces coordonn�es qui suivent sont en pixel, pas en twips
    
    ' R�cup�re la position du RTB par rapport � l'�cran
    Call GetWindowRect(RTB.hwnd, Moi)
    
    ' R�cup�re la position du curseur par rapport � l'�cran
    Call GetCursorPos(Souris)
    
    ' Si le curseur est en dehors de notre RTB, on ne fait rien
    ' (en fait pas tr�s utile puisque on n'arrive ici que si la
    '  souris bouge dans le controle RTB, mais si on met ce code
    '  dans une Sub, �a servira)
    If Not (Souris.x >= Moi.Left And Souris.x <= Moi.Right And _
            Souris.y >= Moi.Top And Souris.y <= Moi.Bottom) Then
                Me.Caption = "Souris hors du texte"
                Exit Sub
    End If
    
    ' Coordonn�es du caract�re sous la souris
    Caract.x = Souris.x - Moi.Left
    Caract.y = Souris.y - Moi.Top
    
    ' Recherche le caract�re qui correspond
    PosCar = SendMessage(RTB.hwnd, EM_CHARFROMPOS, ByVal 0, ByVal Caract)
    If PosCar = 0 Then
        Me.Caption = "Souris hors du texte"
        Exit Sub
    End If
    Temp = Mid(RTB.Text, PosCar, 1)
    'Me.Caption = "PosCar = " & CStr(PosCar) & " - Caract�re = """ & Temp & """"

    '-- Recherche le mot complet et le s�lectionne
    RTB.SelStart = PosCar
    ' Cherche l'espace avant
    Temp = "( ;.,/=+-)" & vbCr & vbLf    ' caract�res qui d�limitent un mot
    RTB.Span Temp, False, True  ' devant
    RTB.Span Temp, True, True   ' derri�re
    Me.Caption = RTB.SelText
    
End Sub

