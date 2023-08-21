
Sub MacroTest()
'
' MacroTest Macro
'
    wsName = "OpTimeAggregate"

    Dim sht As Worksheet
    For Each sht In Worksheets
    If sht.Name = wsName Then
        Worksheets(wsName).Delete
    End If
    Next

    title_row = 3
    end_row = 0
    
    ActiveWindow.SmallScroll ToRight:=-2
    ActiveWindow.SmallScroll Down:=0
    Rows("3:3").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    ' Dim wb, ws, title_row
    Dim wss
    Set wb = Application.ActiveWorkbook
    Set ws = wb.Worksheets.Add(before:=Sheets(1)) 'Create a new sheet
    ws.Name = wsName
    MsgBox wsName & " table is created!"
    ' write_row = ws.UsedRange.Rows.Count + 1
    write_row = title_row

    sheet_row = Worksheets("Latest data from BPR").UsedRange.Rows.Count
    Worksheets("Latest data from BPR").Rows(title_row & ":" & sheet_row - end_row).Copy ws.Range("A" & write_row)
    
    Dim SourceData As Worksheet: Set SourceData = Worksheets(1)

    SourceDataLstr = SourceData.Range("A" & Rows.Count).End(xlUp).Row 'Find the lastrow in the Source Data Sheet
    ' MappingLstr = Mapping.Range("A" & Rows.Count).End(xlUp).Row 'Find the lastrow in the Mapping Sheet

    ' searchTexts = Split("AMS,Operate", ",")
    searchTexts = ExtractColumnValuesToArray("MS Engagements", 3)

    ' Dim core_op_team As String
    ' core_op_team = Split("Nick McEwen,Thomas Gross,Shyam Kumar,David Ang,Kobe Xu,Siva Anbalagan,Ryan Cruz- PDC,Ma. Jesusa Cruz- PDC", ",")
    core_op_team = ExtractColumnValuesToArray("Core Operate Team", 1)
    
    With SourceData

        For i = title_row To SourceDataLstr
            MatterFieldValue = .Cells(i, "E").Value
            ClientFieldValue = .Cells(i, "D").Value
            Debug.Print "MatterFieldValue: " & MatterFieldValue
            Debug.Print "ClientFieldValue: " & ClientFieldValue
            .Cells(i, "W").Value = .Cells(i, "C")   ' Copy staff names
            .Cells(i, "AA").Value = .Cells(i, "F")   ' Copy matter desc detail
            If i > title_row Then
            For ti = LBound(searchTexts) To UBound(searchTexts)
                Debug.Print "searchTexts: " & searchTexts(ti)
                bFind = False
                If InStr(searchTexts(ti), MatterFieldValue) > 0 And InStr(searchTexts(ti), ClientFieldValue) > 0 Then
                    .Cells(i, "Z").Value = .Cells(i, "G")   ' Set operate hours
                    bFind = True
                    Exit For
                End If
            Next ti
            
            If bFind = False Then
                 .Cells(i, "Y").Value = .Cells(i, "G")   ' Set non operate hours
            End If
            
            Dim StaffFieldValue As String
            StaffFieldValue = .Cells(i, "C").Value
          
            StaffFieldValue = Replace(StaffFieldValue, ", ", ",")
            StaffFieldValue = Replace(StaffFieldValue, ",", " ")

            For mi = LBound(core_op_team) To UBound(core_op_team)
                Dim tar_name_parts() As String
                Dim name_parts() As String
                
                tar_name_parts() = Split(core_op_team(mi), " ")
                name_parts() = Split(StaffFieldValue, " ")
                'Debug.Print "myArray is of type: " & TypeName(tar_name_parts)
                'Debug.Print "myArray is of type: " & VarType(tar_name_parts)
                
                ' Check if arrays are equal (unordered)
                ' result = True

                result = AreArraysEqualUnordered1(name_parts, tar_name_parts)
                ' If InStr(StaffFieldValue, core_op_team(mi)) > 0 Then
                If result = True Then
                    .Cells(i, "X").Value = "Y"    ' from core team
                    Exit For
                Else
                    .Cells(i, "X").Value = "N"    ' not from core team
                End If
            Next mi

            Else
                .Cells(i, "W").Value = "Staff Name"
                ' .Cells(i, "X").Value = "Non Operate Hours"
                .Cells(i, "X").Value = "Core Team"
                .Cells(i, "Y").Value = "Non Operate Hours"
                .Cells(i, "Z").Value = "Operate Hours"
                .Cells(i, "AA").Value = "Matter Desc"
            End If

        Next i

    End With
    ' ' add buffer
    ' Set startCell = ws.Range("A3")
    ' Set endCell = startCell.End(xlDown)
    ' Set endCell = endCell.End(xlToRight)
    ' lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' lastColumn = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    ' Dim cellAddress As String
    ' cellAddress = endCell.Address
    ' Debug.Print endCell.Address

    ' Dim lRow As Integer
    ' Dim lCol As Integer
    ' Dim rng As Range
    ' Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ws.Range(startCell, Cells(lastRow, lastColumn)))
    
    ' ' ' add pilot table
    ' Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("Z3"), TableName:="MyPivotTable")

    ' With pt
    '     .PivotFields("Staff Name_NotOP").Orientation = xlRowField
    '     .AddDataField .PivotFields("Chargable"), "Total Sales", xlSum
    ' End With
End Sub
