VERSION 5.00
Begin VB.Form FrmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
    If TxtTemps.Text = "" Then
        MsgBox "Le temps entre chaque image prise par la webcam doit �tre plus d'au moins une seconde", vbCritical, "Erreur"
    ElseIf Int(TxtTemps.Text) = 0 Then
        MsgBox "Le temps entre chaque image prise par la webcam doit �tre plus d'au moins une seconde", vbCritical, "Erreur"
    ElseIf TxtNbDes.Text = "" Then
        MsgBox "Il faut lancer au moins un d�", vbCritical, "Erreur"
    ElseIf Int(TxtNbDes.Text) > 4 Then
        MsgBox "3 d�s maximum", vbCritical, "Erreur"
    ElseIf Int(TxtNbDes.Text) = 0 Then
        MsgBox "Il faut lancer au moins un d�", vbCritical, "Erreur"
    ElseIf TxtScore.Text = "" Then
        MsgBox "Le score � atteindre doit �tre supp�rieur � z�ro", vbCritical, "Erreur"
    ElseIf Int(TxtScore.Text) > (Int(TxtNbDes.Text) * 6) * 2 / 3 Then
        MsgBox "Le score � atteindre doit �tre inf�rieur ou �gal aux 2/3 de la valeur maxiumum atteignable par tout les d�s (ici :" & (Int(TxtNbDes.Text) * 6) * 2 / 3 & ")", vbCritical, "Erreur"
    ElseIf Int(TxtScore.Text) = 0 Then
        MsgBox "Le score � atteindre doit �tre supp�rieur � 0", vbCritical, "Erreur"
    Else
        FrmMain.Show
        FrmMain.ClkWebcam.Enabled = True
        FrmOption.Hide
    End If
End Sub
