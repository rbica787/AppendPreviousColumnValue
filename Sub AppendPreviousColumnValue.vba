Sub AppendPreviousColumnValue()
    Dim cell As Range
    Dim prevCellValue As String

    ' Loop through each cell in the selected range
    For Each cell In Selection
        If Not IsEmpty(cell.Offset(0, -1)) Then ' Check if there is a value in the previous column
            prevCellValue = cell.Offset(0, -1).Value
            
            ' Check if the previous cell value is a single-digit number
            If IsNumeric(prevCellValue) And Len(prevCellValue) = 1 Then
                prevCellValue = "0" & prevCellValue ' Add leading zero if it's a single digit
            End If
            
            ' Append previous column value to the end
            cell.Value = cell.Value & prevCellValue
        End If
    Next cell
End Sub

