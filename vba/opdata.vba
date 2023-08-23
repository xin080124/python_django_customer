
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
    ' write_row = ws.UsedRange.Rows.Count + 1
    write_row = title_row

    sheet_row = Worksheets("Latest data from BPR").UsedRange.Rows.Count
    Worksheets("Latest data from BPR").Rows(title_row & ":" & sheet_row - end_row).Copy ws.Range("A" & write_row)
    
    Dim SourceData As Worksheet: Set SourceData = Worksheets(1)

    SourceDataLstr = SourceData.Range("A" & Rows.Count).End(xlUp).Row 'Find the lastrow in the Source Data Sheet
    ' MappingLstr = Mapping.Range("A" & Rows.Count).End(xlUp).Row 'Find the lastrow in the Mapping Sheet

    ' searchTexts = Split("AMS,Operate", ",")
    OPClients = ExtractColumnValuesToArray("MS Engagements", 1)
    OPMatters = ExtractColumnValuesToArray("MS Engagements", 2)
    Leaves = ExtractColumnValuesToArray("MS Engagements", 4)

    ' Dim core_op_team As String
    core_op_team = ExtractColumnValuesToArray("Core Operate Team", 1)

    ' Dim myDictionary() As Variant

    Dim myDictionary(1 To 11, 1 To 2) As Variant

    ' ReDim myDictionary(dictLowerBound To dictUpperBound, 1 To 2)
    
    ' Add key-value pairs
    ' .Cells(i, "W").Value = "Staff Name"
    ' .Cells(i, "X").Value = "Core Team"
    ' .Cells(i, "Y").Value = "Non Operate Hours"
    ' .Cells(i, "Z").Value = "Operate Hours"
    ' .Cells(i, "AA").Value = "Matter Desc"
    myDictionary(1, 1) = "Staff Name Copy"
    myDictionary(1, 2) = "W"
    
    myDictionary(2, 1) = "Core Team"
    myDictionary(2, 2) = "X"
    
    myDictionary(3, 1) = "Other Engagements"
    myDictionary(3, 2) = "Y"

    myDictionary(10, 1) = "Leave Hours"
    myDictionary(10, 2) = "Z"

    myDictionary(4, 1) = "Operate Hours"
    myDictionary(4, 2) = "AA"

    myDictionary(5, 1) = "Client & Matter Desc"
    myDictionary(5, 2) = "AB"

    myDictionary(6, 1) = "Chargable"
    myDictionary(6, 2) = "F"

    myDictionary(7, 1) = "Matter Desc"
    myDictionary(7, 2) = "E"

    myDictionary(8, 1) = "Client Sort Name"
    myDictionary(8, 2) = "D"

    myDictionary(9, 1) = "Staff Name"
    myDictionary(9, 2) = "C"

    myDictionary(11, 1) = "Total Hours"
    myDictionary(11, 2) = "O"
    
    With SourceData

        For i = title_row To SourceDataLstr
            MatterFieldValue = .Cells(i, GetValue(myDictionary, "Matter Desc")).Value
            ClientFieldValue = .Cells(i, GetValue(myDictionary, "Client Sort Name")).Value
            Debug.Print "MatterFieldValue: " & MatterFieldValue
            Debug.Print "ClientFieldValue: " & ClientFieldValue
            .Cells(i, GetValue(myDictionary, "Staff Name Copy")).Value = .Cells(i, GetValue(myDictionary, "Staff Name"))   ' Copy staff names
            .Cells(i, GetValue(myDictionary, "Client & Matter Desc")).Value = .Cells(i, GetValue(myDictionary, "Client Sort Name")) & " " & .Cells(i, GetValue(myDictionary, "Matter Desc"))
   ' Copy matter desc detail
            If i > title_row Then
            For ti = LBound(OPClients) To UBound(OPClients)
                Debug.Print "LCase(OPClients(ti)): " & LCase(OPClients(ti))
                Debug.Print "LCase(ClientFieldValue): " & LCase(ClientFieldValue)
                bFind = False
                Debug.Print "InStr(LCase(MatterFieldValue), LCase(OPMatters(ti))) : " & InStr(LCase(MatterFieldValue), LCase(OPMatters(ti)))
                Debug.Print "InStr(LCase(OPClients(ti)), LCase(ClientFieldValue)) : " & InStr(LCase(OPClients(ti)), LCase(ClientFieldValue))
                If Len(ClientFieldValue) > 0 And Len(OPMatters(ti)) > 0 And Len(OPClients(ti)) > 0 And InStr(LCase(MatterFieldValue), LCase(OPMatters(ti))) > 0 And InStr(LCase(ClientFieldValue), LCase(OPClients(ti))) > 0 Then
                    .Cells(i, GetValue(myDictionary, "Operate Hours")).Value = .Cells(i, GetValue(myDictionary, "Chargable"))   ' Set operate hours
                    bFind = True
                    Exit For
                Else 
                    For li = LBound(Leaves) To UBound(Leaves)
                        If InStr(LCase(MatterFieldValue), LCase(Leaves(li))) > 0 Then
                            .Cells(i, GetValue(myDictionary, "Leave Hours")).Value = .Cells(i, GetValue(myDictionary, "Total Hours"))   ' Set leave hours
                            bFind = True
                            Exit For
                        End If
                    Next li
                End If
            Next ti
            
            If bFind = False Then
                 .Cells(i, GetValue(myDictionary, "Other Engagements")).Value = .Cells(i, GetValue(myDictionary, "Chargable"))   ' Set other engagements hours
            End If
            
            Dim StaffFieldValue As String
            StaffFieldValue = .Cells(i, GetValue(myDictionary, "Staff Name")).Value
          
            StaffFieldValue = Replace(StaffFieldValue, ", ", ",")
            StaffFieldValue = Replace(StaffFieldValue, ",", " ")

            For mi = LBound(core_op_team) To UBound(core_op_team)
                Dim tar_name_parts() As String
                Dim name_parts() As String
                
                
                core_op_team(mi) = Replace(core_op_team(mi), ", ", ",")
                core_op_team(mi) = Replace(core_op_team(mi), ",", " ")
                tar_name_parts() = Split(core_op_team(mi), " ")
                name_parts() = Split(StaffFieldValue, " ")
                'Debug.Print "myArray is of type: " & TypeName(tar_name_parts)
                'Debug.Print "myArray is of type: " & VarType(tar_name_parts)
                
                ' Check if arrays are equal (unordered)
                ' result = True

                result = AreArraysEqualUnordered1(name_parts, tar_name_parts)
                ' If InStr(StaffFieldValue, core_op_team(mi)) > 0 Then
                If result = True Then
                    .Cells(i, GetValue(myDictionary, "Core Team")).Value = "Y"    ' from core team
                    Exit For
                Else
                    .Cells(i, GetValue(myDictionary, "Core Team")).Value = "N"    ' not from core team
                End If
            Next mi

            Else
                .Cells(i, GetValue(myDictionary, "Staff Name Copy")).Value = "Staff Name Copy"
                .Cells(i, GetValue(myDictionary, "Core Team")).Value = "Core Team"
                .Cells(i, GetValue(myDictionary, "Other Engagements")).Value = "Other Engagements"
                .Cells(i, GetValue(myDictionary, "Leave Hours")).Value = "Leave Hours"
                .Cells(i, GetValue(myDictionary, "Operate Hours")).Value = "Operate Hours"
                .Cells(i, GetValue(myDictionary, "Client & Matter Desc")).Value = "Client & Matter Desc"
                ' .Cells(i, "W").Value = "Staff Name"
            End If

        Next i

    End With
    MsgBox wsName & " table is created!"
    ' ' add buffer
    ' Set startCell = ws.Range("W3")
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
    ' Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("AC3"), TableName:="MyPivotTable")

    ' With pt
    '     .PivotFields("Staff Name Copy").Orientation = xlRowField
    '     .AddDataField .PivotFields("Non Operate Hours"), "Total Sales", xlSum
    ' End With
End Sub

' Retrieve a value
Function GetValue(myDictionary As Variant, key As String) As Variant
    Dim i As Integer
    ' For i = LBound(myDictionary, 1) To UBound(myDictionary, 1)
    For i = 1 To 11
        If myDictionary(i, 1) = key Then
            GetValue = myDictionary(i, 2)
            Exit Function
        End If
    Next i
    ' Return Empty if the key doesn't exist
    GetValue = Empty
End Function





