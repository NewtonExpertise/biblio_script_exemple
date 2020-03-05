Sub pivottable()
'
' pivottable Macro
' pivottablepivottable
'

'
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Feuil1!R1C1:R16C5", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Feuil8!R3C1", TableName:="Tableau croisé dynamique5", _
        DefaultVersion:=xlPivotTableVersion15
    Sheets("Feuil8").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Tableau croisé dynamique5").PivotFields( _
        "individus")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tableau croisé dynamique5").AddDataField ActiveSheet. _
        PivotTables("Tableau croisé dynamique5").PivotFields("pizza"), "Somme de pizza" _
        , xlSum
End Sub



