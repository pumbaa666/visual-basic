VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim fichier As Long
    Dim strVar As String
    Dim fileName As String
    fileName = "D:\Projects\MyProjects\Shutdown_VB\pense.txt"
    'fileName = ".\pense.txt"
    'fileName = CurDir & "\pense.txt"
    
    fichier = FreeFile
    
    If FileLen(fileName) > 0 Then
        Open fileName For Input As #fichier
            Line Input #fichier, strVar
            If strVar <> "" Then
                If MsgBox(strVar & vbCrLf & "Ok pour éteindre, Annuler pour rester sur Windaube et annuler le compte à rebours", vbOKCancel) = vbOK Then
                    ShutDown_Windaube
                End If
            Else
                ShutDown_Windaube
            End If
            
        Close #fichier
    Else
        ShutDown_Windaube
    End If
    
    Unload Me
    
End Sub

Private Sub ShutDown_Windaube()
    Shell "shutdown -s -t 60"
    
End Sub
