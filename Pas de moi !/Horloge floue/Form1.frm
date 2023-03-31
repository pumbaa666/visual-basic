VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   345
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   6165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type IconeTray
'tim As Timer 'Timer de l'icone Tray
cbSize As Long 'Taille de l'icône (en octets)
hWnd As Long 'Handle de la fenêtre chargée de recevoir les messages envoyés lors des évènements sur l'icône (clics, doubles-clics...)
uID As Long 'Identificateur de l'icône
uFlags As Long
uCallbackMessage As Long 'Messages à renvoyer
hIcon As Long 'Handle de l'icône
szTip As String * 64 'Texte à mettre dans la bulle d'aide
End Type

Dim IconeT As IconeTray


'Constantes nécessaires pour la gestion de l'icône dans le systray
Private Const AJOUT = &H0
Private Const MODIF = &H1
Private Const SUPPRIME = &H2
Private Const MOUSEMOVE = &H200
Private Const MESSAGE = &H1
Private Const Icone = &H2
Private Const TIP = &H4
'Constante pour la gestion des évènement souris sur l'icône du systray
Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205

'API nécessaire pour afficher l'icône du systray
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconeTray) As Boolean
Private Sub Form_Load()
    Load Form2  'Chargement de la form2 qui contient le menu PopUpmenu / Dans si le popupmenu était dans cette form on aurais un cadre autour de l'horloge
    Me.Left = Screen.Width - Label1.Width - 1500    'Mise en place de l'appli pour ne pas cacher les boutons Agrandir, Réduire, Fermer
    Me.Top = 0
    Timer1_Timer    'Force le déclenchement du timer pour afficher quelque chose

    'Préparation de la variable IconeT pour le systray
    IconeT.cbSize = Len(IconeT) 'Taille de l'icône en octet
    IconeT.hWnd = Me.hWnd 'Handle de l'application (pour qu'elle reçoive les messages envoyés lors d'un clic, double-clic...
    IconeT.uID = 1& 'Identificateur de l'icône
    IconeT.uFlags = Icone Or TIP Or MESSAGE
    IconeT.uCallbackMessage = MOUSEMOVE 'Renvoyer les messages concernant l'action de la souris
    IconeT.hIcon = Me.Icon 'Mettre en icône l'image qui est dans le contrôle "Image1"
    IconeT.szTip = "Horloge Floue" & Chr$(0) 'Texte de la bulle d'aide

    'Appel de la fonction pour mettre l'icône dans le système tray
    Shell_NotifyIcon AJOUT, IconeT
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static msg As Long  'Gestion des clic souris sur l'icône
    msg = X / Screen.TwipsPerPixelX
    Select Case msg 'Différentes possibilité d'action
        Case DOUBLE_CLICK_GAUCHE: 'mettez
        Case BOUTON_GAUCHE_POUSSE: 'ce
        Case BOUTON_GAUCHE_LEVE: 'que
            PopupMenu Form2.MnuMain
        Case DOUBLE_CLICK_DROIT: 'vous
        Case BOUTON_DROIT_POUSSE: 'voudrez
        Case BOUTON_DROIT_LEVE: 'qu'il se passe
            PopupMenu Form2.MnuMain
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Désaffichage de l'icône dans le systray
    IconeT.cbSize = Len(IconeT)
    IconeT.hWnd = Me.hWnd
    IconeT.uID = 1&
    Shell_NotifyIcon SUPPRIME, IconeT
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Form2.MnuMain
    End If
End Sub

Private Sub Timer1_Timer()
    'Déclaration des variables
    Dim H, M, A As Integer
    Dim Heure, heureb, Min, Minb As String

    A = GetSetting(App.Title, "Settings", "Level", 0)   'Récupère le mode de fonctionnement de l'horloge (Floue, un peu, beaucoup, etc...)
    H = DatePart("h", Now)  'Récupère l'heure
    M = DatePart("n", Now)  'Récupère les minutes

    If A = 1 Then
        Select Case M   'Selon les minutes actuelle
            Case 53 To 59
                H = H + 1    'H prend 1 heure pour arondir l'heure actuelle
            Case 0 To 7
                Min = ""    'Min est vide
            Case 8 To 22
                Min = "et quart"
            Case 23 To 37
                Min = "et demi"
            Case 38 To 52
                H = H + 1  'H+1 pour faire le : 18 heure moins le quart par exemple
                Min = "moins le quart"
        End Select
        If H = 24 Then H = 0
    End If

    'Les heures sont mises en lettres également / Placer ici car comme on en a besoin, cela prend moins de place

    On Error Resume Next
        If Val(GetSetting(App.Title, "Settings", "Desk", "1")) = 0 Then
            Me.Visible = False
        Else
            Me.Visible = True
        End If
    On Error GoTo 0

    If Val(GetSetting(App.Title, "Settings", "Format", "1")) = 1 Then
        Heure = Choose(H + 1, "Minuit", _
            "Une", _
            "Deux", _
            "Trois", _
            "Quatre", _
            "Cinq", _
            "Six", _
            "Sept", _
            "Huit", _
            "Neuf", _
            "Dix", _
            "Onze", _
            "Midi", _
            "Treize", _
            "Quatorze", _
            "Quinze", _
            "Seize", _
            "Dix-sept", _
            "Dix-Huit", _
            "Dix-Neuf", _
            "Vingt", _
            "Vingt et une", _
            "Vingt deux", _
            "Vingt trois")
    Else
        Heure = Choose(H + 1, "Minuit", _
            "Une", _
            "Deux", _
            "Trois", _
            "Quatre", _
            "Cinq", _
            "Six", _
            "Sept", _
            "Huit", _
            "Neuf", _
            "Dix", _
            "Onze", _
            "Midi", _
            "Une", _
            "Deux", _
            "Trois", _
            "Quatre", _
            "Cinq", _
            "Six", _
            "Sept", _
            "Huit", _
            "Neuf", _
            "Dix", _
            "Onze")
    End If

    If A = 0 Then   'Mode précis
        'Selon les minutes actuelles, on applique les minutes en toutes lettres dans Minb (M+1) car si M = 0 alors choose le voit comme M = -1
        Minb = Choose(M + 1, "Zéro", "Une", "Deux", "Trois", "Quatre", "Cinq", "Six", "Sept", "Huit", "Neuf", "Dix", _
        "Onze", "Douze", "Treize", "Quatorze", "Quinze", "Seize", "Dix-Sept", "Dix-Huit", "Dix-Neuf", "Vingt", _
        "Vingt && Une", "Vingt Deux", "Vingt Trois", "Vingt Quatre", "Vingt Cinq", "Vingt Six", "Vingt Sept", "Vingt Huit", "Vingt Neuf", "Trente", _
        "Trente && Une", "Trente Deux", "Trente Trois", "Trente Quatre", "Trente Cinq", "Trente Six", "Trente Sept", "Trente Huit", "Trente Neuf", "Quarante", _
        "Quarante && Une", "Quarante Deux", "Quarante Trois", "Quarante Quatre", "Quarante Cinq", "Quarante Six", "Quarante Sept", "Quarante Huit", "Quarante Neuf", "Cinquante", _
        "Cinquante && Une", "Cinquante Deux", "Cinquante Trois", "Cinquante Quatre", "Cinquante Cinq", "Cinquante Six", "Cinquante Sept", "Cinquante Huit", "Cinquante Neuf")

        If Heure = "Minuit" Or Heure = "Midi" Then  'Si il est midi ou minuit il ne faut pas que heure apparaîsse derrière, exemple : midi heure
            Label1.Caption = Heure & " " & Minb
        Else
            Label1.Caption = Heure & " Heure et " & Minb & " Minutes"
        End If
    ElseIf A = 1 Then   'Mode Un peu floue
        If Heure = "Minuit" Or Heure = "Midi" Then
            Label1.Caption = Heure & " " & Min
        Else
            Label1.Caption = Heure & " Heure " & Min
        End If
    ElseIf A = 2 Then   'Mode beaucoup floue
        Select Case H   'Selon l'heure actuelle
            Case 4 To 10    'Entre 4 et 10 h du matin
                Label1.Caption = "C'est le matin"
            Case 11 To 13
                Label1.Caption = "C'est le midi"
            Case 14 To 18
                Label1.Caption = "C'est l'après midi"
            Case 19 To 23
                Label1.Caption = "C'est le soir"
            Case 0 To 3
                Label1.Caption = "C'est la nuit"
        End Select
    ElseIf A = 3 Then   'Mode A la folie
        Label1.Caption = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    ElseIf A = 4 Then   'Mode dans le vague
        Select Case Format(Now, "dddd") 'Selon le jour on détermine où on en ai dans la semaine
            Case "lundi"
                Label1.Caption = "C'est le début de la semaine"
            Case "mardi"
                Label1.Caption = "C'est le début de la semaine"
            Case "mercredi"
                Label1.Caption = "C'est la mi-semaine"
            Case "jeudi"
                Label1.Caption = "C'est la mi-semaine"
            Case "vendredi"
                Label1.Caption = "C'est bientôt le week-end"
            Case "samedi"
                Label1.Caption = "C'est le week-end !!!"
            Case "dimanche"
                Label1.Caption = "C'est le week-end !!!"
        End Select
    End If

    'Mise à jour de l'icône systray
    If Val(GetSetting(App.Title, "Settings", "Systray", "3")) = 0 Then
        IconeT.szTip = Format(Now, "hh:mm")
    ElseIf Val(GetSetting(App.Title, "Settings", "Systray", "3")) = 1 Then
        IconeT.szTip = Format(Now, "dddd d mmm yyyy") & "   " & Format(Now, "hh:mm")
    ElseIf Val(GetSetting(App.Title, "Settings", "Systray", "3")) = 2 Then
        IconeT.szTip = Format(Now, "dddd d mmm yyyy")
    ElseIf Val(GetSetting(App.Title, "Settings", "Systray", "3")) = 3 Then
        IconeT.szTip = Label1.Caption
    End If

    Shell_NotifyIcon MODIF, IconeT  'Mise à jour de l'icône tray
End Sub
