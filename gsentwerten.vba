Sub FindAndEnterText()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim searchValue As String
    Dim compareValue As String
    Dim cell As Range
    
    Dim foundCounter As Integer
    foundCounter = 0
    
    ' Set references to the worksheets
    Set ws1 = ThisWorkbook.Sheets("Planung23")
    Set ws2 = ThisWorkbook.Sheets("GSALL")
    
    For i = 821 To 850

        ' Get the value to search for from cell A1 in Sheet1
        ' searchValue = ws1.Range("A1").Value
        searchValue = ws1.Cells(i, 8).Value
        
        If searchValue <> "" And searchValue <> "bar" Then
        
            ' Loop through each cell in column A of Sheet2
            For Each cell In ws2.Range("B:B")
                compareValue = cell.Value
                ' If cell.Value = "" Then Exit For ' Exit loop if empty cell encountered
                If cell.Value = searchValue Then
                    ' If value is found, enter "Found" in column B ofd the same row.
                    cell.Offset(0, 3).Value = "09/02/2023"
                    foundCounter = foundCounter + 1
                    Exit For ' Exit loop if value is found
                End If
            Next cell
        
        End If
    Next i
    
    MsgBox "Found: " & foundCounter
End Sub


