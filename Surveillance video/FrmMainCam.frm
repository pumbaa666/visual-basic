VERSION 5.00
Object = "{102225D5-EA25-11D3-886E-00105A154A4D}#1.0#0"; "VPortal2.dll"
Begin VB.Form FrmMainCam 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transmission en cours"
   ClientHeight    =   4245
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ClkTrans 
      Interval        =   250
      Left            =   600
      Top             =   3600
   End
   Begin VPORTAL2LibCtl.VideoPortal VideoPortal1 
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "FrmMainCam.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMainCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClkTrans_Timer() ' Fais clignoter les petits points dans le titre de la Form
Static vNbPoint As Integer
    frmMainCam.Caption = frmMainCam.Caption & "."
    vNbPoint = vNbPoint + 1
    If vNbPoint = 10 Then
        vNbPoint = 0
        frmMainCam.Caption = "Transmission en cours"
    End If
End Sub

Private Sub Form_Load()
    VideoPortal1.PrepareControl "QCSDK_VBDEMO", "HKEY_LOCAL_MACHINE\Software\Logitech\QCSDK_VBDEMO", 0
    
    ' Essaie de connecter une caméra
    If VideoPortal1.ConnectCamera2() = 0 Then
        MsgBox "Impossible de connecter une caméra", vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' Si une caméra est trouvée ça enclanche la prévisualisation de l'image
    VideoPortal1.EnablePreview = 1
End Sub
