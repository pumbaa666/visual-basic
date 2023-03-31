VERSION 5.00
Begin VB.Form frmSplash 
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMSPL~1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   6120
         Top             =   3600
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright 2004"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Société SDan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   4
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licence accordée à Sébastian DANCOT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "SDan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   5
         Top             =   720
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Form_Load()
Show
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblProductName.Caption = "Le Questionnaire"
End Sub
Public Sub Timer1_Timer()
frmmenu.Show
Unload Me
End Sub
