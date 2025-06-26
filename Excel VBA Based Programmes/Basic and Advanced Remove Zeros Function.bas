Attribute VB_Name = "Module1"
'This function removes any zeros from a given array based on a specific row/column. The output is the same array excluding zeros from the specific row/column

'rng is the input data (e.g. a range, array, table, etc)
'dim_num is the row/col to filter the zeros by (1 if a single row/col)
'dims is the dimension to search through (0 for col, 1 for row)
Function remove_zeros(rng As Variant, dim_num As Long, dims As Integer) As Variant

Dim data As Variant
Dim result() As Variant '() indicate array, (1:10) is a fixed array, () is an adjustable array
Dim temp() As Variant
Dim count As Long
Dim i As Long, j As Long
Dim row_count As Long, col_count As Long

'checks if the input data (i.e. rng) is an array and if it is not then converts it to an array
If Not IsArray(rng) Then
    data = rng.Value
Else:
    data = rng
End If

count = 0
'extracts the number of rows and columns from the array/range
row_count = UBound(data, 1)
col_count = UBound(data, 2)
'redefines the dimensions for the array "result" as the size of the input data (i.e. data)
ReDim result(1 To row_count, 1 To col_count) 'redefines the index of the range (i.e. 1:2 to 1:3) while preserve the current values in the array (Preserve)

'checks if all cells of a specified column are not equal to zero adds all rows that are not to a new temporary array
'it then adjusts the dimensions of the final output, transfering the information from the temporary arrary
If dims = 0 Then
    For i = 1 To row_count
        If data(i, dim_num) <> 0 Then
            count = count + 1
            For j = 1 To col_count
                result(count, j) = data(i, j)
            Next j
        End If
    Next i
    ReDim temp(1 To count, 1 To col_count)
    For i = 1 To count
        For j = 1 To col_count
            temp(i, j) = result(i, j)
        Next j
    Next i
    
'checks if all cells of a specified row are not equal to zero adds all columns that are not to a new temporary array
'it then adjusts the dimensions of the final output, transfering the information from the temporary arrary
Else
    For i = 1 To col_count
        If data(dim_num, i) <> 0 Then
            count = count + 1
            For j = 1 To row_count
                result(j, count) = data(j, i)
            Next j
        End If
    Next i
    ReDim temp(1 To row_count, 1 To count)
    For j = 1 To row_count
        For i = 1 To count
            temp(j, i) = result(j, i)
        Next i
    Next j
End If

remove_zeros = temp

End Function
'Removes zeros from a 1D range
Function basic_remove_zeros(rng As Range)

Dim result() As Variant '() indicate array, (1:10) is a fixed array, () is an adjustable array
Dim count As Long
Dim cell As Range

count = 0
For Each cell In rng
    If cell.Value <> 0 Then
        count = count + 1
        ReDim Preserve result(1 To count) 'redefines the index of the range (i.e. 1:2 to 1:3) while preserve the current values in the array (Preserve)
        result(count) = cell.Value
    End If
Next cell

basic_remove_zeros = result

End Function
