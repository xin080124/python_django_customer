Sub CreatePilotTable()
'
' CreatePilotTable Macro
'

'
    ' Application.CutCopyMode = False
    
    wsName = "Pilot"

    Dim sht As Worksheet
    For Each sht In Worksheets
    If sht.Name = wsName Then
        Worksheets(wsName).Delete
    End If
    Next
    
    Dim wss
    Set wb = Application.ActiveWorkbook
    Set ws = wb.Worksheets.Add(before:=Sheets(1)) 'Create a new sheet
    ws.Name = wsName
    
    ' Sheets.Add
    ' ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    '     "OpTimeAggregate!R3C23:R28C27", Version:=8).CreatePivotTable _
    '     TableDestination:="Sheet2!R1C1", TableName:="PivotTable2", DefaultVersion _
    '     :=8
    ' Sheets("Sheet2").Select
    ' Cells(1, 1).Select
    ' ActiveWorkbook.ShowPivotTableFieldList = True
    ' ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ' ActiveChart.SetSourceData Source:=Range("Sheet2!$A$1:$C$18")
    ' ActiveSheet.Shapes("Chart 1").IncrementLeft 260
    ' ActiveSheet.Shapes("Chart 1").IncrementTop 15
    ' ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
    '     PivotTable.PivotFields("Staff Name Copy"), "Count of Staff Name Copy", xlCount
    ' With ActiveChart.PivotLayout.PivotTable.PivotFields("Staff Name Copy")
    '     .Orientation = xlRowField
    '     .Position = 1
    ' End With
    ' With ActiveChart.PivotLayout.PivotTable.PivotFields("Core Team")
    '     .Orientation = xlRowField
    '     .Position = 1
    ' End With
    ' ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
    '     PivotTable.PivotFields("Non Operate Hours"), "Sum of Non Operate Hours", xlSum
    ' ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
    '     PivotTable.PivotFields("Operate Hours"), "Count of Operate Hours", xlCount
    ' ActiveSheet.Shapes("Chart 1").IncrementLeft 288
    ' ActiveSheet.Shapes("Chart 1").IncrementTop 291
    ' ActiveWindow.Zoom = 167
    ' ActiveChart.PlotArea.Select
    ' Selection.Left = 17.529
    ' Selection.Top = 0
    ' ActiveChart.ChartArea.Select
    ' ActiveSheet.Shapes("Chart 1").IncrementLeft -290.4192125984
    ' ActiveSheet.Shapes("Chart 1").IncrementTop -53.2934645669


    Set startCell = Worksheets("OpTimeAggregate").Range("A3")
    Set endCell = startCell.End(xlDown)
    Set endCell = endCell.End(xlToRight)
    lastRow = Worksheets("OpTimeAggregate").Cells(Worksheets("OpTimeAggregate").Rows.Count, 1).End(xlUp).Row
    lastColumn = Worksheets("OpTimeAggregate").Cells(3, Worksheets("OpTimeAggregate").Columns.Count).End(xlToLeft).Column

    ' Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ws.Range(startCell, Cells(lastRow, lastColumn)))

    ' Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets("OpTimeAggregate").Range(startCell, Cells(lastRow, lastColumn)))

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets("OpTimeAggregate").Range("A3:C5"))

    ' ' add pilot table
    ' Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A1"), TableName:="MyPivotTable")

    ' With pt
    '     .PivotFields("Staff Name Copy").Orientation = xlRowField
    '     .AddDataField .PivotFields("Non Operate Hours"), "Total Sales", xlSum
    ' End With
    MsgBox wsName & " Pilot table is created!"

End Sub
