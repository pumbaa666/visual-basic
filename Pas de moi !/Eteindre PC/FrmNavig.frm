VERSION 5.00
Begin VB.Form FrmNavig 
   Caption         =   "Micro PC - Shut Down"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   Icon            =   "FrmNavig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptAction 
      Caption         =   "Redémarrage"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Tag             =   "1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.OptionButton OptAction 
      Caption         =   "Arrêter"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Tag             =   "2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton OptAction 
      Caption         =   "Fermer la session"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Tag             =   "3"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Que voulez-vous faire ?"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actuellement"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   3615
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Il est"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdValider 
      Caption         =   "Sauvegarder ces paramètres"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sélectionner l'heure"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   3615
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   13
         Text            =   "00:00:00"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Déconnexion à : hh:mm:ss"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sélectionner le jour"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
      Begin VB.OptionButton OptJour 
         Caption         =   "jeudi"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   8
         Tag             =   "4"
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2640
         Top             =   960
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   3120
         Top             =   960
      End
      Begin VB.OptionButton OptJour 
         Caption         =   "dimanche"
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   11
         Tag             =   "7"
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptJour 
         Caption         =   "samedi"
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   10
         Tag             =   "6"
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptJour 
         Caption         =   "vendredi"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   9
         Tag             =   "5"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptJour 
         Caption         =   "mercredi"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptJour 
         Caption         =   "mardi"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Tag             =   "2"
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OptJour 
         Caption         =   "lundi"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Tag             =   "1"
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmNavig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'*      Programme développé par Emmanuel Prevot (JPeman)   *
'*                                                         *
'*  Programme de fermeture de PC à partir d'un fichier     *
'*  Setup.log généré lors de la sauvegarde des paramètres  *
'*  de redemarrage, de fermeture de session ou d'arret de  *
'*  l'ordinateur. Tout est dans le source !!!              *
'*                                                         *
'***********************************************************

Option Explicit

Dim JourStr
Dim ActionStr


'***********************************************************
'*  Ce début code source utilisé a été trouvé sur vbfrance *
'***********************************************************

'APIs
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

'Constants
Private Const EWX_FORCE As Long = 4

'Types
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type

'Enumerations
Public Enum EnumExitWindows

WE_LOGOFF = 0
WE_SHUTDOWN = 1
WE_REBOOT = 2
WE_POWEROFF = 8

End Enum

'Variables

'Functions and Subs
Private Sub AdjustToken()

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long

hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle

' Get the LUID for shutdown privilege.
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

tkp.PrivilegeCount = 1 ' One privilege to set
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED

' Enable the shutdown privilege in the access token of this process.
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded

End Sub

Public Sub ExitWindows(ByVal l_Command As EnumExitWindows)

AdjustToken

ExitWindowsEx (l_Command Or EWX_FORCE), 0

End Sub

'************************************************************
'*  Ce début code source utilisé a été développé par JPeman *
'************************************************************

'Commande du bouton Sauvegarder les parmaetres
Private Sub CmdValider_Click()
    Ecrire_Fichier
    Timer1.Enabled = True
End Sub


'Fonction de lecture du fichier Setup.log
Public Function Lire_Fichier()
    Const Lecture = 1
    Dim fso, Fichier_TXT, ArrFichier, NomFichier
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Fichier_TXT = fso.OpenTextFile("c:\Shut Down\Setup.Log", Lecture)
    NomFichier = Fichier_TXT.ReadLine
    
    If NomFichier <> "" Then
        ArrFichier = Split(NomFichier, "|")
        'Format de Fichier : redemarrage|lundi|14:00:00
        If ArrFichier(0) = "Redémarrage" Then
            OptAction(0).Value = True
        ElseIf ArrFichier(0) = "Arrêter" Then
            OptAction(2).Value = True
        Else
            OptAction(1).Value = True
        End If
                
        Dim i
        For i = 0 To 6 'nombre de jours - 1
            If ArrFichier(1) = OptJour(i).Caption Then
                OptJour(i).Value = True
            End If
        Next
        
        Text1.Text = ArrFichier(2)
    Else
        ActionStr = "Redémarrage"
        OptAction(0).Value = True
        JourStr = "Lundi"
        OptJour(0).Value = True
        Text1.Text = "12:00:00"
    End If
End Function

'Fonction d'ecriture du fichier Setup.log
Public Function Ecrire_Fichier()
    Dim ChaineStr
    Dim i, j
        
    On Error Resume Next
    'Format de Fichier : redemarrage|lundi|14:00:00
    
    ChaineStr = ""
    ActionStr = ""
    JourStr = ""
    
    For i = 0 To 2
        If OptAction(i) Then
            ActionStr = OptAction(i).Caption
        End If
    Next
    
    If ActionStr = "" Then ActionStr = "Redémarrage":
    
    For i = 0 To 6
        If OptJour(i) Then
            JourStr = OptJour(i).Caption
        End If
    Next
    
    If JourStr = "" Then JourStr = "Lundi":
    
    ChaineStr = ActionStr
    ChaineStr = ChaineStr & "|" & JourStr
    ChaineStr = ChaineStr & "|" & Text1.Text
    
    Open "c:\Shut Down\Setup.log" For Output As 1
    Print #1, ChaineStr
    Close #1
End Function

'Ouverture de l'application
Private Sub Form_Load()
    Lire_Fichier
    Timer1.Enabled = True
End Sub

'Fonction de gestion du clic sur les options de fermetures
Private Sub OptAction_Click(Index As Integer)
    Dim i
    For i = 0 To 2
        If OptAction(i) Then
            ActionStr = OptAction(i).Caption
        End If
    Next
End Sub

'Fonction de gestion du clic sur les jours
Private Sub OptJour_Click(Index As Integer)
    Dim i
    For i = 0 To 6
        If OptJour(i) Then
            JourStr = OptJour(i).Caption
        End If
    Next
End Sub

'Fonction de gestion du timer (temps) pour heure choisie
Private Sub Timer1_Timer()
    If Text1.Text = Text2.Text Then 'heure actuelle = heure choisie
        'Attention aux majuscules (gros piege) car le programme prend la casse !
        If WeekdayName(Weekday(Day(Date))) = CStr(JourStr) Then 'jour actuel = jour choisi
            If OptAction(0).Value = True Then 'redemarrage sélectionné
                ExitWindows WE_REBOOT
            ElseIf OptAction(1).Value = True Then 'fermeture de session sélectionnée
                ExitWindows WE_LOGOFF
            ElseIf OptAction(2).Value = True Then 'arret sélectionné
                ExitWindows WE_SHUTDOWN
            End If
        End If
    End If
End Sub

'Fonction de gestion du timer 2 pour heure actuelle (temps)
Private Sub Timer2_Timer()
    Text2.Text = Time
End Sub
