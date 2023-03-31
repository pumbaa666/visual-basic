VERSION 5.00
Begin VB.Form FrmChrono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chronogramme"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LblFin 
      Caption         =   "1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label LblArrivee 
      Caption         =   "0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   400
   End
   Begin VB.Shape ShpPlace 
      Height          =   495
      Index           =   0
      Left            =   360
      Top             =   240
      Width           =   1000
   End
   Begin VB.Label LblNom 
      Alignment       =   2  'Center
      Caption         =   "Nom"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "FrmChrono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
