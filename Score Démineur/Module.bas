Attribute VB_Name = "Module1"
Public tTabAbsLoic() As Currency
Public tTabAbsAlderic() As Currency
Public vJoueur1 As String
Public vJoueur2 As String

Function Chargement2()
Dim vCount As Integer
    '**** Victoires ****'
    FrmMain.LGradY(0).Y1 = FrmMain.AxeX.Y1 - (3480 - 900) / 13
    FrmMain.LGradY(0).Y2 = FrmMain.AxeX.Y1 - (3480 - 900) / 13
    FrmMain.LGradY(0).X1 = FrmMain.LGradY(0).X1 + 50
    FrmMain.LGradY(0).X2 = FrmMain.LGradY(0).X2 - 50
    FrmMain.LblAxeY(0).Top = FrmMain.AxeX.Y1 - (3480 - 900) / 13 - 120
    FrmMain.LblAxeY(0).Caption = "2"

    '**** Score Absolu ****'
    FrmMain.AbsGradY(0).Y1 = FrmMain.AbsAxeX.Y1 - (3480 - 900) / 13
    FrmMain.AbsGradY(0).Y2 = FrmMain.AbsAxeX.Y1 - (3480 - 900) / 13
    FrmMain.AbsGradY(0).X1 = FrmMain.AbsGradY(0).X1 + 50
    FrmMain.AbsGradY(0).X2 = FrmMain.AbsGradY(0).X2 - 50
    FrmMain.LblAbsAxeY(0).Top = FrmMain.AbsAxeX.Y1 - (3480 - 900) / 13 - 120
    If Int(FrmMain.LblAbsoluAlderic.Caption) > Int(FrmMain.LblAbsoluAlderic.Caption) Then
        FrmMain.LblAbsAxeY(0).Caption = Int(Int(FrmMain.LblAbsoluAlderic.Caption) / 13 / 2)
    Else
        FrmMain.LblAbsAxeY(0).Caption = Int(Int(FrmMain.LblAbsoluLoic.Caption) / 13 / 2)
    End If

    '**** Victoires ****'
    For vCount = 1 To 26
        Load FrmMain.LGradY(vCount)
        FrmMain.LGradY(vCount).Y1 = FrmMain.AxeX.Y1 - ((3480 - 900) / 26) * vCount
        FrmMain.LGradY(vCount).Y2 = FrmMain.AxeX.Y1 - ((3480 - 900) / 26) * vCount
        
        If vCount Mod 2 = 0 Then
            FrmMain.LGradY(vCount).X1 = FrmMain.LGradY(vCount).X1 - 50
            FrmMain.LGradY(vCount).X2 = FrmMain.LGradY(vCount).X2 + 50
        End If
        FrmMain.LGradY(vCount).Visible = True

        If (vCount - 2) Mod 4 = 0 Then
            Load FrmMain.LblAxeY(vCount)
            FrmMain.LblAxeY(vCount).Caption = vCount
            FrmMain.LblAxeY(vCount).Top = FrmMain.LGradY(vCount).Y1 - 120
            FrmMain.LblAxeY(vCount).Visible = True
        End If
    Next

    '**** Score absolu ****'
    For vCount = -8 To 26
        Load FrmMain.AbsGradY(vCount + 9)
        FrmMain.AbsGradY(vCount + 9).Y1 = FrmMain.AbsAxeX.Y1 - ((3480 - 900) / 26) * vCount
        FrmMain.AbsGradY(vCount + 9).Y2 = FrmMain.AbsAxeX.Y1 - ((3480 - 900) / 26) * vCount
        If Abs(vCount + 9) Mod 2 = 0 Then
            FrmMain.AbsGradY(vCount + 9).X1 = FrmMain.AbsGradY(vCount + 9).X1 - 50
            FrmMain.AbsGradY(vCount + 9).X2 = FrmMain.AbsGradY(vCount + 9).X2 + 50
        End If
        FrmMain.AbsGradY(vCount + 9).Visible = True

        If (vCount + 9 - 2) Mod 4 = 0 Then
            Load FrmMain.LblAbsAxeY(vCount + 9)
            If Int(FrmMain.LblAbsoluAlderic.Caption) > Int(FrmMain.LblAbsoluAlderic.Caption) Then
                FrmMain.LblAbsAxeY(vCount + 9).Caption = Int(Int(FrmMain.LblAbsoluAlderic.Caption) / 13 * vCount / 2)
            Else
                FrmMain.LblAbsAxeY(vCount + 9).Caption = Int(Int(FrmMain.LblAbsoluLoic.Caption) / 13 * vCount / 2)
            End If
            FrmMain.LblAbsAxeY(vCount + 9).Top = FrmMain.AbsGradY(vCount + 9).Y1 - 120
            FrmMain.LblAbsAxeY(vCount + 9).Visible = True
        End If
    Next
End Function

Function Chargement()
Dim vLecture As String

Dim vNbMatch As Integer
Static vOldNbMatch As Integer

Dim vVictLoic As Integer
Dim vVictAlderic As Integer

Dim vOldVictLoic As Integer
Dim vOldVictAlderic As Integer

Dim vNbVictLoic As Integer
Dim vNbVictAlderic As Integer

Dim vVictConsLoic As Integer
Dim vVictConsAlderic As Integer

Dim vRecVictConsLoic As Integer
Dim vRecVictConsAlderic As Integer

Dim vFrite As Integer
Dim vFriteLoic As Integer
Dim vFriteAlderic As Integer
Dim vFriteLoicStr As String
Dim vFriteAldericStr As String

Dim vParfait As Integer

Static vChargement As Boolean

Dim vOldScoreAbsLoic As Currency
Dim vOldScoreAbsAlderic As Currency

Dim vNbMines As Integer

    FrmMain.Show
    ReDim tTabAbsLoic(0)
    ReDim tTabAbsAlderic(0)
    vJoueur1 = Mid(FrmFichier.File1.FileName, 1, InStr(FrmFichier.File1.FileName, "-") - 1)
    vJoueur2 = Mid(FrmFichier.File1.FileName, InStr(FrmFichier.File1.FileName, "-") + 1, Len(FrmFichier.File1.FileName) - Len(vJoueur1) - 5)
    FrmMain.LblJoueur1 = vJoueur1
    FrmMain.LblJoueur2 = vJoueur2
    FrmMain.Lbl2Joueur1 = vJoueur1
    FrmMain.Lbl2Joueur2 = vJoueur2
    FrmMain.Lbl3Joueur1 = vJoueur1
    FrmMain.Lbl3Joueur2 = vJoueur2
    Open FrmFichier.TxtFichier.Text For Input As #1
    Do
        '**** Lecture des scores ****'
        Line Input #1, vLecture

        If vLecture = "26-25" Or vLecture = "25-26" Then
            vParfait = vParfait + 1
        End If

        vOldVictAlderic = vVictAlderic
        vOldVictLoic = vVictLoic

        vVictAlderic = Right(vLecture, 2)
        vVictLoic = Left(vLecture, 2)

        vNbMines = vNbMines + vVictAlderic + vVictLoic

        If vNbMatch > 0 Then
            vOldScoreAbsLoic = tTabAbsLoic(vNbMatch - 1)
            vOldScoreAbsAlderic = tTabAbsAlderic(vNbMatch - 1)
        End If
        
        ReDim Preserve tTabAbsLoic(vNbMatch)
        ReDim Preserve tTabAbsAlderic(vNbMatch)
        tTabAbsLoic(vNbMatch) = vNbVictLoic - vNbVictAlderic + 5 * vRecVictConsLoic + vFriteLoic / 4
        tTabAbsAlderic(vNbMatch) = vNbVictAlderic - vNbVictLoic + 5 * vRecVictConsAlderic + vFriteAlderic / 4

        '************* Graphique **************'
        With FrmMain
            '***** Création *****'
            If vChargement = False Then
                If vNbMatch > 1 Then
                    Load .LGraphLoic(vNbMatch - 1)
                    Load .LGraphAlderic(vNbMatch - 1)
                    Load .LGradX(vNbMatch - 1)
                    Load .LblAxeX(vNbMatch - 1)

                    Load .GraphAbsLoic(vNbMatch - 1)
                    Load .GraphAbsAlderic(vNbMatch - 1)
                    Load .AbsGradX(vNbMatch - 1)
                    Load .LblAbsAxeX(vNbMatch - 1)
                End If
            Else
                If vNbMatch = vOldNbMatch Then
                    Load .LGraphLoic(vNbMatch - 1)
                    Load .LGraphAlderic(vNbMatch - 1)
                End If
            End If
            If vNbMatch <> 0 Then
                '********** Victoires **********'
                .LGraphLoic(vNbMatch - 1).X1 = .AxeY.X1 + vNbMatch * 150
                .LGraphLoic(vNbMatch - 1).X2 = .AxeY.X1 + (vNbMatch + 1) * 150
                .LGraphLoic(vNbMatch - 1).Y1 = .AxeX.Y1 - 100 * vOldVictLoic
                .LGraphLoic(vNbMatch - 1).Y2 = .AxeX.Y1 - 100 * vVictLoic

                .LGraphLoic(vNbMatch - 1).Visible = True
                .LGraphLoic(vNbMatch - 1).BorderColor = &HFF&

                .LGradX(vNbMatch - 1).X1 = .LGraphLoic(vNbMatch - 1).X1
                .LGradX(vNbMatch - 1).X2 = .LGraphLoic(vNbMatch - 1).X1

                If (vNbMatch - 1) Mod 2 = 0 Then
                    .LGradX(vNbMatch - 1).Y1 = .AxeX.Y1 - 100
                    .LGradX(vNbMatch - 1).Y2 = .AxeX.Y1 + 110

                    .LblAxeX(vNbMatch - 1).Left = .LGradX(vNbMatch - 1).X1 - 240
                    .LblAxeX(vNbMatch - 1).Caption = vNbMatch
                    .LblAxeX(vNbMatch - 1).Visible = True
                Else
                    .LGradX(vNbMatch - 1).Y1 = .AxeX.Y1 - 45
                    .LGradX(vNbMatch - 1).Y2 = .AxeX.Y1 + 55
                End If
                .LGradX(vNbMatch - 1).Visible = True

                .LGraphAlderic(vNbMatch - 1).X1 = .AxeY.X1 + vNbMatch * 150
                .LGraphAlderic(vNbMatch - 1).X2 = .AxeY.X1 + (vNbMatch + 1) * 150
                .LGraphAlderic(vNbMatch - 1).Y1 = .AxeX.Y1 - 100 * vOldVictAlderic
                .LGraphAlderic(vNbMatch - 1).Y2 = .AxeX.Y1 - 100 * vVictAlderic

                .LGraphAlderic(vNbMatch - 1).Visible = True
                .LGraphAlderic(vNbMatch - 1).BorderColor = &HFF0000

                '********** Score absolu **********'
                .GraphAbsLoic(vNbMatch - 1).X1 = .AbsAxeY.X1 + vNbMatch * 150
                .GraphAbsLoic(vNbMatch - 1).X2 = .AbsAxeY.X1 + (vNbMatch + 1) * 150
                .GraphAbsLoic(vNbMatch - 1).Y1 = .AbsAxeX.Y1 - 52 * vOldScoreAbsLoic
                .GraphAbsLoic(vNbMatch - 1).Y2 = .AbsAxeX.Y1 - 52 * tTabAbsLoic(vNbMatch)

                .GraphAbsLoic(vNbMatch - 1).Visible = True
                .GraphAbsLoic(vNbMatch - 1).BorderColor = &HFF&

                .AbsGradX(vNbMatch - 1).X1 = .GraphAbsLoic(vNbMatch - 1).X1
                .AbsGradX(vNbMatch - 1).X2 = .GraphAbsLoic(vNbMatch - 1).X1

                If (vNbMatch - 1) Mod 2 = 0 Then
                    .AbsGradX(vNbMatch - 1).Y1 = .AbsAxeX.Y1 - 100
                    .AbsGradX(vNbMatch - 1).Y2 = .AbsAxeX.Y1 + 110

                    .LblAbsAxeX(vNbMatch - 1).Left = .AbsGradX(vNbMatch - 1).X1 - 70
                    .LblAbsAxeX(vNbMatch - 1).Caption = vNbMatch
                    .LblAbsAxeX(vNbMatch - 1).Visible = True
                Else
                    .AbsGradX(vNbMatch - 1).Y1 = .AbsAxeX.Y1 - 45
                    .AbsGradX(vNbMatch - 1).Y2 = .AbsAxeX.Y1 + 55
                End If
                .AbsGradX(vNbMatch - 1).Visible = True

                .GraphAbsAlderic(vNbMatch - 1).X1 = .AbsAxeY.X1 + vNbMatch * 150
                .GraphAbsAlderic(vNbMatch - 1).X2 = .AbsAxeY.X1 + (vNbMatch + 1) * 150
                .GraphAbsAlderic(vNbMatch - 1).Y1 = .AbsAxeX.Y1 - 47 * vOldScoreAbsAlderic
                .GraphAbsAlderic(vNbMatch - 1).Y2 = .AbsAxeX.Y1 - 47 * tTabAbsAlderic(vNbMatch)

                .GraphAbsAlderic(vNbMatch - 1).Visible = True
                .GraphAbsAlderic(vNbMatch - 1).BorderColor = &HFF0000
            End If
        End With
        
        '****** Qui a gagné ?!? *****'
        If vVictLoic > vVictAlderic Then
            vNbVictLoic = vNbVictLoic + 1

            '***** Victoires consécutives *****'
            vVictConsLoic = vVictConsLoic + 1
            vVictConsAlderic = 0
            If vVictConsLoic > vRecVictConsLoic Then
                vRecVictConsLoic = vVictConsLoic
            End If

            '************** Fritée ************'
            vFrite = vVictLoic - vVictAlderic
            If vFrite > vFriteLoic Then
                vFriteLoic = vFrite
                vFriteLoicStr = vLecture
            End If
        ElseIf vVictLoic <> vVictAlderic Then
            vNbVictAlderic = vNbVictAlderic + 1

            '***** Victoires consécutives *****'
            vVictConsAlderic = vVictConsAlderic + 1
            vVictConsLoic = 0
            If vVictConsAlderic > vRecVictConsAlderic Then
                vRecVictConsAlderic = vVictConsAlderic
            End If

            '************** Fritée ************'
            vFrite = vVictAlderic - vVictLoic
            If vFrite > vFriteAlderic Then
                vFriteAlderic = vFrite
                vFriteAldericStr = vLecture
            End If
        End If

        '*** Affichage des scores ***'
        FrmMain.LstLoic.AddItem vVictLoic
        FrmMain.LstAlderic.AddItem vVictAlderic

        vNbMatch = vNbMatch + 1
    Loop While (Not EOF(1))
    Close #1

    '*** Affichage des victoires ****'
    With FrmMain
        '**** Victoires ****'
        .AxeX.X2 = .LGraphAlderic(vNbMatch - 2).X2 + 500
        .Width = .AxeX.X2 + 500

        If vNbMatch > vOldNbMatch Then
            Load .LGradX(vNbMatch - 1)
            Load .AbsGradX(vNbMatch - 1)
        End If
        .LGradX(vNbMatch - 1).X1 = .AxeY.X1 + vNbMatch * 150
        .LGradX(vNbMatch - 1).X2 = .AxeY.X1 + vNbMatch * 150

        '**** Score absolu ****'
        .AbsAxeX.X2 = .GraphAbsAlderic(vNbMatch - 2).X2 + 500
        .AbsGradX(vNbMatch - 1).X1 = .AbsAxeY.X1 + vNbMatch * 150
        .AbsGradX(vNbMatch - 1).X2 = .AbsAxeY.X1 + vNbMatch * 150
        If (vNbMatch - 1) Mod 2 = 0 Then
            If vNbMatch > vOldNbMatch Then
                Load .LblAxeX(vNbMatch - 1)
                Load .LblAbsAxeX(vNbMatch - 1)
            End If
            .LblAxeX(vNbMatch - 1).Left = .LGradX(vNbMatch - 1).X1 - 240
            .LblAxeX(vNbMatch - 1).Caption = vNbMatch
            .LblAxeX(vNbMatch - 1).Visible = True

            .LblAbsAxeX(vNbMatch - 1).Left = .AbsGradX(vNbMatch - 1).X1 - 70
            .LblAbsAxeX(vNbMatch - 1).Caption = vNbMatch
            .LblAbsAxeX(vNbMatch - 1).Visible = True
        Else
            .AbsGradX(vNbMatch - 1).Y1 = .AbsAxeX.Y1 + 50
            .AbsGradX(vNbMatch - 1).Y2 = .AbsAxeX.Y1 - 50

            .AbsGradX(vNbMatch - 1).Y1 = .AbsAxeX.Y1 + 50
            .AbsGradX(vNbMatch - 1).Y2 = .AbsAxeX.Y1 - 50
        End If

        .LGradX(vNbMatch - 1).Visible = True
        .AbsGradX(vNbMatch - 1).Visible = True
        
        .LblVictAlderic.Caption = vNbVictAlderic
        .LblVictLoic.Caption = vNbVictLoic

        .LblPourcentAlderic.Caption = Left(Trim(Str(vNbVictAlderic / vNbMatch * 100)), 5)
        .LblPourcentLoic.Caption = Left(Trim(Str(vNbVictLoic / vNbMatch * 100)), 5)

        .LblVictConsAlderic.Caption = vVictConsAlderic
        .LblVictConsLoic.Caption = vVictConsLoic

        .LblRecVictConsAlderic.Caption = vRecVictConsAlderic
        .LblRecVictConsLoic.Caption = vRecVictConsLoic

        .LblFriteLoic.Caption = vFriteLoicStr
        .LblFriteAlderic.Caption = Right(vFriteAldericStr, 2) & "-" & Left(vFriteAldericStr, 2)

        .LblParfait.Caption = vParfait

        .LblNbMatch.Caption = vNbMatch

        .LblMine.Caption = vNbMines

        .LblAbsoluAlderic.Caption = tTabAbsAlderic(vNbMatch - 1)
        .LblAbsoluLoic.Caption = tTabAbsLoic(vNbMatch - 1)
    End With
    vChargement = True
    vOldNbMatch = vNbMatch
    Chargement2
    If FrmMain.Width < 8500 Then
        FrmMain.Width = 8500
    End If
End Function
