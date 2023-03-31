VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6000
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "scalpweb@hotmail.com"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Form Modifi�e � l'aide d'APIs WIndows."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Si vous avez le moindre probl�me de compr�hension :
'// scalpweb@hotmail.com
'// Ce serait avec plaisir que je vous r�pondrais !

Private Sub Form_Load()

'// Ici, on va appliquer les r�gions de fa�on
'// � cr�er des r�gions invisibles.
'// Les APIs sont d�clar�es dans le module.

'// Variables n�c�ssaires :
Dim rgnCercle As Long
Dim rgnBarre As Long
Dim rgnTrou As Long
Dim rgnFinale As Long
'// Elles sont de type "Long" car les fonctions
'// des APIs renvoies des valeurs de type "Long".

'// On initalise les variables :
'// elle permettent de cr�er des regions invisibles
'// de forme rectangulaires ou circulaires.
rgnCercle = CreateEllipticRgn(100, 0, 300, 200)
rgnBarre = CreateRectRgn(0, 80, 400, 120)

'// On cr�e la zone principale :
rgnFinale = CreateRectRgn(0, 0, 400, 200)

'// On combine toutes les zones :
'// On utilise RGN_OR comme op�rateur logique.
CombineRgn rgnFinale, rgnCercle, rgnBarre, RGN_OR

'// On associe la r�gion combin�e � la form :
SetWindowRgn Me.hwnd, rgnFinale, True

End Sub
