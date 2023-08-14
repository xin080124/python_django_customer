
' Sub TestArrayEqualityUnordered()
'     Dim arr1() As String
'     Dim arr2() As String
'     Dim result As Boolean
    
'     ' Initialize arrays
'     arr1 = Split("A,B,C,D,E,F", ",")
'     arr2 = Split("F,E,D,C,B,A", ",")
    
'     ' Check if arrays have the same elements (unordered)
'     result = AreArraysEqualUnordered(arr1, arr2)
    
'     ' Display result
'     If result Then
'         MsgBox "The two arrays have the same elements (unordered)."
'     Else
'         MsgBox "The two arrays do not have the same elements (unordered)."
'     End If
' End Sub

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

    sheet_row = Worksheets(3).UsedRange.Rows.Count
    Worksheets(3).Rows(title_row & ":" & sheet_row - end_row).Copy ws.Range("A" & write_row)
    
    Dim SourceData As Worksheet: Set SourceData = Worksheets(1)

    SourceDataLstr = SourceData.Range("A" & Rows.Count).End(xlUp).Row 'Find the lastrow in the Source Data Sheet
    ' MappingLstr = Mapping.Range("A" & Rows.Count).End(xlUp).Row 'Find the lastrow in the Mapping Sheet

    searchTexts = Split("AMS,Operate", ",")
    
     ' Dim core_op_team As String
    core_op_team = Split("NC", ",")
    With SourceData

        For i = title_row To SourceDataLstr
            MatterFieldValue = .Cells(i, "E").Value
            .Cells(i, "W").Value = .Cells(i, "C")
            
            If i > title_row Then
            For ti = LBound(searchTexts) To UBound(searchTexts)
                If InStr(MatterFieldValue, searchTexts(ti)) > 0 Then
                    .Cells(i, "Z").Value = .Cells(i, "G")
                    Exit For
                Else
                    .Cells(i, "X").Value = .Cells(i, "G")
                End If
            Next ti

            Dim StaffFieldValue As String
            StaffFieldValue = .Cells(i, "C").Value
            ' StaffFieldValue = "McEwen, Nick"
            
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
                    .Cells(i, "Y").Value = "Y"
                    Exit For
                Else
                    .Cells(i, "Y").Value = "N"
                End If
            Next mi

            Else
                .Cells(i, "W").Value = "Staff Name"
                .Cells(i, "X").Value = "Non Operate Hours"
                .Cells(i, "Y").Value = "Core Team"
                .Cells(i, "Z").Value = "Operate Hours"
                .Cells(i, "X").Value = "Non Operate Hours"
                .Cells(i, "Y").Value = "Core Team"
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
