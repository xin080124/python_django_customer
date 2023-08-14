
Function AreArraysEqualUnordered1(metaParts() As String, tarParts() As String) As Boolean
    
    Dim matchFlag As Boolean
    For i = LBound(metaParts) To UBound(metaParts)
        matchFlag = False
        For j = LBound(tarParts) To UBound(tarParts)
            If metaParts(i) = tarParts(j) Then
                matchFlag = True
            End If
        Next j
        If matchFlag = False Then
            AreArraysEqualUnordered1 = matchFlag
            Exit Function
        End If
    Next i
    
AreArraysEqualUnordered1 = matchFlag

End Function

Function AreArraysEqualUnordered(arr1() As String, arr2() As String) As Boolean
    Dim col As Collection
    Set col = New Collection
    
    Dim i As Long
    Dim item As Variant
    
    ' Add elements from the first array to the collection
    For i = LBound(arr1) To UBound(arr1)
        col.Add False
    Next i
    
    Dim bFlag As Boolean
    
    ' Check if elements from the second array are in the collection
    For i = LBound(arr1) To UBound(arr1)
        bFlag = False
        If col.item(i) <> False Then
             bFalg = True
        End If
        
        If bFlag <> True Then
            For j = LBound(arr2) To UBound(arr2)
                If arr2(j) = arr1(i) Then
                    bFlag = True
                End If
            Next j
            
            If bFlag = True Then
                col.item(i) = True
            Else
                AreArraysEqualUnordered = False
                Exit Function
            End If
        End If
    Next i

    ' All elements are equal
    AreArraysEqualUnordered = True
End Function

