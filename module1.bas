Sub ClearCellsByStyle()
    Dim rng As Range
    Dim cell As Range
    Dim inputColor As Long
    
    ' Define the background color (e.g., light yellow in RGB format)
    inputColor = RGB(255, 204, 153) ' Equivalent to #FFCC99
    
    ' Loop through each cell in the active sheet
    For Each cell In ActiveSheet.UsedRange
    
        ' Check if the cell has the specified background color and a hyperlink
        If cell.Interior.Color = inputColor And cell.Hyperlinks.Count > 0 Then
            ' Clear the hyperlink
            cell.Hyperlinks.Delete
            
            ' Set the cell style to "Input"
            cell.Style = "Input"
        End If
        
        ' Check if the cell has the style named "Input"
        If cell.Style = "Input" Then
            ' Clear the value of the cell
            cell.ClearContents
        End If
        
    Next cell
End Sub
