VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "� propos de MonApplication"
   ClientHeight    =   6495
   ClientLeft      =   8355
   ClientTop       =   2235
   ClientWidth     =   3990
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4482.964
   ScaleMode       =   0  'User
   ScaleWidth      =   3746.816
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   5760
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2520
      TabIndex        =   0
      Top             =   6120
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Infos syst�me..."
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4425
      Left            =   360
      Picture         =   "frmAbout.frx":030A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   3605.958
      Y1              =   3892.828
      Y2              =   3892.828
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   3605.958
      Y1              =   3892.828
      Y2              =   3892.828
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   3165
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Options de s�curit� des cl�s de base de registres...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Types racines des cl�s de base de registres...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cha�ne termin�e par un caract�re nul Unicode.
Const REG_DWORD = 4                      ' Nombre 32 bits.

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "� propos de " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Essaie d'obtenir le chemin et le nom du programme Infos syst�me dans la base de registre...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Essaie d'obtenir uniquement le chemin du programme Infos syst�me dans la base de registre...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Valide l'existence de la version du fichier 32 bits connu.
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Erreur - Impossible de trouver le fichier...
        Else
            GoTo SysInfoErr
        End If
    ' Erreur - Impossible de trouver l'entr�e de la base de registre...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Les informations syst�me ne sont pas disponibles actuellement", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Compteur de boucle.
    Dim rc As Long                                          ' Code de retour.
    Dim hKey As Long                                        ' Descripteur d'une cl� de base de registres ouverte.
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Type de donn�es d'une cl� de base de registres.
    Dim tmpVal As String                                    ' Stockage temporaire pour une valeur de cl� de base de registres.
    Dim KeyValSize As Long                                  ' Taille de la variable de la cl� de base de registres.
    '------------------------------------------------------------
    ' Ouvre la cl� de base de registres sous la racine cl� {HKEY_LOCAL_MACHINE...}.
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvre la cl� de base de registres.
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' G�re l'erreur...
    
    tmpVal = String$(1024, 0)                             ' Alloue de l'espace pour la variable.
    KeyValSize = 1024                                       ' D�finit la taille de la variable.
    
    '------------------------------------------------------------
    ' Extrait la valeur de la cl� de base de registres...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtient/Cr�e la valeur de la cl�.
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' G�re l'erreur.
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ajoute une cha�ne termin�e par un caract�re nul...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Caract�re nul trouv�, extrait de la cha�ne.
    Else                                                    ' WinNT ne termine pas la cha�ne par un caract�re nul...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Caract�re nul non trouv�, extrait la cha�ne uniquement.
    End If
    '------------------------------------------------------------
    ' D�termine le type de valeur de la cl� pour la conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Recherche les types de donn�es...
    Case REG_SZ                                             ' Type de donn�es cha�ne de la cl� de la base de registres.
        KeyVal = tmpVal                                     ' Copie la valeur de la cha�ne.
    Case REG_DWORD                                          ' Type de donn�es double mot de la cl� de base de registres.
        For i = Len(tmpVal) To 1 Step -1                    ' Convertit chaque bit.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construit la valeur caract�re par caract�re.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertit le mot double en cha�ne.
    End Select
    
    GetKeyValue = True                                      ' Retour avec succ�s.
    rc = RegCloseKey(hKey)                                  ' Ferme la cl� de base de registres
    Exit Function                                           ' Quitte.
    
GetKeyError:      ' R�initialise apr�s qu'une erreur s'est produite...
    KeyVal = ""                                             ' Affecte une cha�ne vide � la valeur de retour.
    GetKeyValue = False                                     ' Retour avec �chec.
    rc = RegCloseKey(hKey)                                  ' Ferme la cl� de base de registres.
End Function

Private Sub lblDescription_Click()

End Sub

