' Formatter.bas - Module VBA pour l'Add-in Excel
' Automatisation du formatage des tableaux

Attribute VB_Name = "Formatter"

Sub FormatTableau()
    Dim ws As Worksheet
    Dim rng As Range
    Dim tbl As ListObject
    
    ' Définir la feuille active
    Set ws = ActiveSheet
    
    ' Vérifier si une plage est sélectionnée
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Veuillez sélectionner une plage de données.", vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Convertir la sélection en tableau s'il n'est pas déjà formaté
    On Error Resume Next
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Impossible de créer un tableau. Vérifiez la sélection.", vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Appliquer un style au tableau
    tbl.TableStyle = "TableStyleMedium9"
    
    ' Ajuster les colonnes
    rng.Columns.AutoFit
    
    ' Mettre en gras la première ligne
    rng.Rows(1).Font.Bold = True
    
    ' Supprimer les doublons (sur la première colonne)
    tbl.Range.RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' Ajouter une colonne de total si applicable
    If tbl.ListColumns.Count > 1 Then
        tbl.ListColumns.Add
        tbl.ListColumns(tbl.ListColumns.Count).Name = "Total"
        tbl.ListColumns(tbl.ListColumns.Count).DataBodyRange.Formula = "=SUM(A2:A" & rng.Rows.Count & ")"
    End If
    
    MsgBox "Le tableau a été formaté avec succès !", vbInformation, "Succès"
End Sub
