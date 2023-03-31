VERSION 5.00
Begin VB.Form FrmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mes DVD's - Patience..."
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Lecture du fichier Excel, veuillez patienter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "FrmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
