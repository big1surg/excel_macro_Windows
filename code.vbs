Sub SetFloors()
    Dim floorRangeSelect As Range
    Dim floorTextSelect As Range
    Dim floorRange As Range
    Dim floorText As Range
    Set floorText = ActiveSheet.Range("E2:E850")
    Set floorRange = ActiveSheet.Range("H2:H850")
    
    For Each floorTextSelect In floorText.Cells
        If InStr(1, floorTextSelect.Value, "CSI") Then
            floorTextSelect.Offset(0, 3).Value = "CSI"
        ElseIf InStr(1, floorTextSelect.Value, "Eulalia") Then
            floorTextSelect.Offset(0, 3).Value = "Eulalia"
        ElseIf InStr(1, floorTextSelect.Value, "/1") Or InStr(1, floorTextSelect.Value, "/1 -") Or InStr(1, floorTextSelect.Value, "ITS") Then
            floorTextSelect.Offset(0, 3).Value = 1
        ElseIf InStr(1, floorTextSelect.Value, "/2 -") Or InStr(1, floorTextSelect.Value, "/2") Then
            floorTextSelect.Offset(0, 3).Value = 2
        ElseIf InStr(1, floorTextSelect.Value, "/3 -") Or InStr(1, floorTextSelect.Value, "/3") Then
            floorTextSelect.Offset(0, 3).Value = 3
        ElseIf InStr(1, floorTextSelect.Value, "/4 -") Then
            floorTextSelect.Offset(0, 3).Value = 4
        ElseIf InStr(1, floorTextSelect.Value, "/5 -") Then
            floorTextSelect.Offset(0, 3).Value = 5
        ElseIf InStr(1, floorTextSelect.Value, "/6 -") Then
            floorTextSelect.Offset(0, 3).Value = 6
        ElseIf InStr(1, floorTextSelect.Value, "/7 -") Then
            floorTextSelect.Offset(0, 3).Value = 7
        ElseIf InStr(1, floorTextSelect.Value, "/8 -") Or InStr(1, floorTextSelect.Value, "TEL") Or InStr(1, floorTextSelect.Value, "/8") Then
            floorTextSelect.Offset(0, 3).Value = 8
        ElseIf InStr(1, floorTextSelect.Value, "/B") Then
            floorTextSelect.Offset(0, 3).Value = "Basement"
        ElseIf InStr(1, floorTextSelect.Value, "/R") Then
            floorTextSelect.Offset(0, 3).Value = "Roof"
        Else
            floorTextSelect.Offset(0, 3).Value = ""
        End If
    Next
      '  floorRangeSelect.FormulaR1C1 = ""
    'Range("H2:H850").Select
    
   ' Range("H2").Select
    'ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(""1"",RC[-3])),1,0)"
    'Range("H3").Select
    pivotInformation
End Sub

Sub DeleteNonWindows7()
    ActiveSheet.Range("H1").Value = "Floor"
    Columns("A:H").Select
    Selection.AutoFilter
    
    ActiveSheet.Range("$A$1:$H$1222").AutoFilter Field:=7, Criteria1:=Array( _
        "aix_version_5.2", "cisco_ios_version_12_2_55_se", "Linux", _
        "Microsoft Windows 10 Enterprise", "SuSE(Linux)", "Ubuntu(Linux)", "vmnix-x86", _
        "Windows 2000", "Windows 2003", "Windows 2003 R2", "Windows 2008", _
        "Windows 2008 R2", "Windows 2012 R2", "Windows 2012 Standard", "="), Operator:= _
        xlFilterValues
    Range("A9:G1222").Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$H$835").AutoFilter Field:=7
    
    SetFloors
End Sub



Sub pivotInformation()

    Columns("A:H").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet3!R1C1:R1048576C8", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet5!R3C1", TableName:="PivotTable2", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("Sheet5").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Floor")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("IP name(*Computer)")
        .Orientation = xlRowField
        .Position = 2
    End With
End Sub

