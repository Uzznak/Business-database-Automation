Sub CreerOngletsParLigne()

    Dim wsSource As Worksheet
    Dim wsSommaire As Worksheet
    Dim wsNew As Worksheet
    Dim ws As Worksheet
    Dim maForme As Shape
    Dim shp As Shape

    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim lastRowSommaire As Long
    Dim lastRowNew As Long
    Dim nomFeuille As String
    Dim cellTarget As Range
    Dim maxFeuille As Long
    
    Dim c As Range

    Set wsSource = Worksheets("suivi")
    Set wsSommaire = Worksheets("sommaire")
    
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Vérification des feuilles
    If wsSource Is Nothing Then
        MsgBox "Feuille 'suivi' introuvable."
        Exit Sub
    End If

    If wsSommaire Is Nothing Then
        MsgBox "Feuille 'sommaire' introuvable."
        Exit Sub
    End If

    ' Dernières lignes
    lastRow = wsSource.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastRowSommaire = wsSommaire.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' il faut s'assurer que le fichier de suivi contient quelque chose
    If lastRow < 2 Then
        MsgBox "Aucune donnée trouvée dans le suivi."
        Exit Sub
    End If

    ' Nettoyage des liens
    wsSommaire.Range("H2:H" & wsSommaire.Rows.Count).ClearContents
    
    
    

    ' **********************************************************
    ' SUPPRESSION DES ONGLETS EN TROP
    

    maxFeuille = lastRowSommaire - 1

    For Each ws In Worksheets
    
        If ws.Name Like "Ligne_*" Then
    
            If Val(Replace(ws.Name, "Ligne_", "")) > maxFeuille Then
    
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
    
            End If
    
        End If
    
    Next ws
    
    ' *********************************************************
    ' FEUILLES ANCIENNES ISSUES DU SUIVI

    j = 1

    For i = 2 To lastRow

        nomFeuille = "Ligne_" & j

        Set wsNew = Nothing

        On Error Resume Next
        Set wsNew = Worksheets(nomFeuille)
        On Error GoTo 0

        If wsNew Is Nothing Then

            Set wsNew = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            wsNew.Name = nomFeuille

        Else

            If WorksheetFunction.CountA(wsNew.Range("A2:I100")) = 0 Then
        
                wsNew.Cells.Clear
        
                For Each shp In wsNew.Shapes
                    shp.Delete
                Next shp
        
            Else
        
                ' Mise à jour des entêtes des fiches déjà existantes
        
                For Each c In wsNew.Range("A1:J1")
        
                    If InStr(1, c.Value, "Échéance", vbTextCompare) > 0 Then
                        c.Value = "Échéance action"
                    End If
        
                    If InStr(1, c.Value, "Statut", vbTextCompare) > 0 Then
                        c.Value = "Statut action"
                    End If
        
                Next c
        
                wsNew.Cells(1, "J").Value = "Date dernière modification"
                
                
                ' Type d'action
                With wsNew.Range("A2:A100")
                
                    .Validation.Delete
                
                    .Validation.Add _
                        Type:=xlValidateList, _
                        Formula1:="=TypeAction"
                                        
                End With
                
                ' Criticité
                With wsNew.Range("B2:B100")
                
                    .Validation.Delete
                
                    .Validation.Add _
                        Type:=xlValidateList, _
                        Formula1:="=Criticite"
                
                End With
                
                ' Priorité
              '  With wsNew.Range("F2:F100")
                
                  '  .Validation.Delete
                
                    '.Validation.Add _
                     '   Type:=xlValidateList, _
                      '  Formula1:="=Priorite"
                
              '  End With
                
                ' Statut action
                With wsNew.Range("H2:H100")
                
                    .Validation.Delete
                
                    .Validation.Add _
                        Type:=xlValidateList, _
                        Formula1:="=StatutAction"
                
                End With
                
                
                ' Type d'action
                With wsNew.Range("A2:A100")
                
                    .Validation.Delete
                
                    .Validation.Add _
                        Type:=xlValidateList, _
                        Formula1:="=TypeAction"
                
                End With
                
                ' Criticité
                With wsNew.Range("B2:B100")
                
                    .Validation.Delete
                
                    .Validation.Add _
                        Type:=xlValidateList, _
                        Formula1:="=Criticite"
                End With
                
                ' Priorité
                With wsNew.Range("F2:F100")
                
                    .Validation.Delete
                
                    .Validation.Add _
                        Type:=xlValidateList, _
                        Formula1:="=Priorite"
                
                End With
                
                
                For Each shp In wsNew.Shapes

                    On Error Resume Next
                
                    shp.TextFrame.Characters.Text = Replace(shp.TextFrame.Characters.Text, "?", "")
                
                    On Error GoTo 0
                
                Next shp
                
                
                            
                GoTo LienSuivant
                
                
        
            End If
        
        End If

        ' Type d'action
        With wsNew.Range("A2:A100")
        
            .Validation.Delete
        
            .Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=TypeAction"
        
        End With
        
        ' Criticité
        With wsNew.Range("B2:B100")
        
            .Validation.Delete
        
            .Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=Criticite"
        
        End With
        
        ' Priorité
        With wsNew.Range("F2:F100")
        
            .Validation.Delete
        
            .Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=Priorite"
        
        End With
        
        ' Statut action
        With wsNew.Range("H2:H100")
        
            .Validation.Delete
        
            .Validation.Add _
                Type:=xlValidateList, _
                Formula1:="=StatutAction"
        
        End With

        ' Entêtes
        wsSource.Range("G1:O1").Copy
        wsNew.Range("A1").PasteSpecial xlPasteValues
      
        ' Données
        wsSource.Range("G" & i & ":O" & i).Copy
        wsNew.Range("A2").PasteSpecial xlPasteValues
        
        ' Type d'action
        With wsNew.Range("A2:A100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, _
                Formula1:="=TypeAction"
        End With
        
        ' Criticité
        With wsNew.Range("B2:B100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, _
                Formula1:="=Criticite"
        End With
        
        ' Priorité
        With wsNew.Range("F2:F100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, _
                Formula1:="=Priorite"
        End With
        
        ' Statut action
        With wsNew.Range("H2:H100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, _
                Formula1:="=StatutAction"
        End With

          
        ' Renommer certains entêtes dans les fiches individuelles
        For Each c In wsNew.Range("A1:J1")

            If InStr(1, c.Value, "Échéance", vbTextCompare) > 0 Then
                c.Value = "Échéance action"
            End If
        
            If InStr(1, c.Value, "Statut", vbTextCompare) > 0 Then
                c.Value = "Statut action"
            End If
        
        Next c
                
        ' Navigation
        wsNew.Cells(1, "I").Value = "Navigation"
        wsNew.Columns("I").ColumnWidth = 35
        
        
        
        wsNew.Cells(1, "J").Value = "Date dernière modification"
        wsNew.Columns("J").ColumnWidth = 22


        'bouton
        lastRowNew = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row
        Set cellTarget = wsNew.Cells(lastRowNew + 2, "I")

        Set maForme = wsNew.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            cellTarget.Left, _
            cellTarget.Top, _
            cellTarget.width * 0.95, _
            25)

        With maForme

            .Left = cellTarget.Left + (cellTarget.width - .width) / 2
            .Top = cellTarget.Top

            .TextFrame.Characters.Text = "Retour Sommaire"
            .OnAction = "AllerSommaire"

            .Fill.ForeColor.RGB = RGB(0, 112, 192)
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .TextFrame.Characters.Font.Bold = True
            .Line.Visible = msoFalse

        End With

LienSuivant:

        wsSommaire.Cells(i, "H").FormulaLocal = _
            "=LIEN_HYPERTEXTE(""#'" & nomFeuille & "'!A2"";""Voir le détail"")"

        j = j + 1

    Next i

    
    
    ' NOUVELLES FICHES CRÉÉES DEPUIS LE SOMMAIRE
    '************************************************
    
    
    
    'changer pour appliquer les menus deroulants aux onglets individuels des anciennes lignes
    'Les fiches déjà générées (Ligne_1 à Ligne_150) ne passent jamais dedans
    For i = lastRow + 1 To lastRowSommaire

        nomFeuille = "Ligne_" & (i - 1)

        Set wsNew = Nothing

        On Error Resume Next
        Set wsNew = Worksheets(nomFeuille)
        On Error GoTo 0

        If wsNew Is Nothing Then

            Set wsNew = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            wsNew.Name = nomFeuille

            ' Entêtes identiques au suivi
            wsSource.Range("G1:O1").Copy
            wsNew.Range("A1").PasteSpecial xlPasteValues
            
            For Each c In wsNew.Range("A1:J1")

                If InStr(1, c.Value, "Échéance", vbTextCompare) > 0 Then
                    c.Value = "Échéance action"
                End If
            
                If InStr(1, c.Value, "Statut", vbTextCompare) > 0 Then
                    c.Value = "Statut action"
                End If
            
            Next c

            ' Ligne vide à compléter
            wsNew.Range("A2:I2").ClearContents
            
            ' Type d'action
            With wsNew.Range("A2:A100")
            
                .Validation.Delete
            
                .Validation.Add _
                    Type:=xlValidateList, _
                    Formula1:="=TypeAction"
            
            End With
            
            ' Criticité
            With wsNew.Range("B2:B100")
            
                .Validation.Delete
            
                .Validation.Add _
                    Type:=xlValidateList, _
                    Formula1:="=Criticite"
            
            End With
            
            ' Priorité
            With wsNew.Range("F2:F100")
            
                .Validation.Delete
            
                .Validation.Add _
                    Type:=xlValidateList, _
                    Formula1:="=Priorite"
            
            End With
            
            ' Statut action
            With wsNew.Range("H2:H100")
            
                .Validation.Delete
            
                .Validation.Add _
                    Type:=xlValidateList, _
                    Formula1:="=StatutAction"
            
            End With
             
            ' Navigation
            wsNew.Cells(1, "I").Value = "Navigation"
            wsNew.Columns("I").ColumnWidth = 35
            
            
            ' colonne date de modification
            wsNew.Cells(1, "J").Value = "Date dernière modification"
            wsNew.Columns("J").ColumnWidth = 28
            wsNew.Range("J:J").NumberFormat = "dd/mm/yyyy hh:mm"

            
            Set cellTarget = wsNew.Range("I3")

            Set maForme = wsNew.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                cellTarget.Left, _
                cellTarget.Top, _
                cellTarget.width * 0.95, _
                25)

            With maForme

                .Left = cellTarget.Left + (cellTarget.width - .width) / 2
                .Top = cellTarget.Top

                .TextFrame.Characters.Text = "Retour Sommaire"
                .OnAction = "AllerSommaire"

                .Fill.ForeColor.RGB = RGB(0, 112, 192)
                .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
                .TextFrame.Characters.Font.Bold = True
                .Line.Visible = msoFalse

            End With

        End If
        
                ' Mise à jour systématique des listes déroulantes
        With wsNew.Range("A2:A100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, Formula1:="=TypeAction"
        End With
        
        With wsNew.Range("B2:B100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, Formula1:="=Criticite"
        End With
        
        With wsNew.Range("F2:F100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, Formula1:="=Priorite"
        End With
        
        With wsNew.Range("H2:H100")
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, Formula1:="=StatutAction"
        End With
                

        wsSommaire.Cells(i, "H").FormulaLocal = _
            "=LIEN_HYPERTEXTE(""#'" & nomFeuille & "'!A2"";""Voir le détail"")"

    Next i

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Onglets et liens générés "

End Sub

