Attribute VB_Name = "Module1"
Sub transfer_data()
    'Define variables
    Dim source_wb As Workbook
    Dim source_sh As Worksheet
    Dim current_sh As Worksheet
    Dim filepath As String
    Dim source_sh_name As String, current_sh_name As String
    Dim list_of_columns As Collection
    Dim col As String
    Dim cell As Range
    Dim i As Integer, j As Integer
    Dim last_column As Integer, last_row As Integer
    Dim header_row As Integer
    Dim header_cell As Range
    Dim col_ind As Variant
    
    'Initialises the list and other variables
    Set list_of_columns = New Collection
    i = 2
    j = 1
    
    'Creates input boxes to enter the variable information
    filepath = InputBox("Enter the full file path of the old workbook:")
    source_sh_name = InputBox("Enter the sheet name from the old workbook:")
    current_sh_name = InputBox("Enter the destination sheet name in this workbook:")
    col = InputBox("Enter the specific column name to be searched:")
    header_row = InputBox("Enter the row that contains the headers:")

    'Prevents the screen from updating when opening the workbook
    Application.ScreenUpdating = False

    'Defines the workbook and sheets in use
    Set source_wb = Workbooks.Open(filepath)
    Set source_sh = source_wb.Sheets(source_sh_name)
    Set current_sh = ThisWorkbook.Sheets(current_sh_name)

    'Find all specified columns within the used range of the header row
    last_column = source_sh.UsedRange.Columns.Count
    For Each header_cell In source_sh.Range(source_sh.Cells(header_row, 1), source_sh.Cells(header_row, last_column))
        If LCase(header_cell.Value) = col Then
            list_of_columns.Add header_cell.Column
        End If
    Next header_cell

    'Check if any of the specified columns were found
    If list_of_columns.Count = 0 Then
        MsgBox "No " & col & " columns found. Exiting."
        GoTo Cleanup
    End If

    'Process each specified column
    For Each col_ind In list_of_columns
        If col_ind >= 3 Then                            'Must be >=3 because the function extracts the two cells to the left of the current cell
            current_sh.Columns(j).ColumnWidth = 9.71    'Sets the column width to 9.71 so the date is visible
            last_row = source_sh.Cells(source_sh.Rows.Count, col_ind).End(xlUp).Row
            For Each cell In source_sh.Range(source_sh.Cells(header_row, col_ind), source_sh.Cells(last_row, col_ind))
                If cell.Value <> "" Then                'If the cell isn't empty extract the cell value + 2 to the left and 1 to the right into their corresponding cells in the current worksheet
                    current_sh.Cells(i, j).Value = source_sh.Cells(cell.Row, col_ind - 2).Value
                    current_sh.Cells(i, j + 1).Value = source_sh.Cells(cell.Row, col_ind - 1).Value
                    current_sh.Cells(i, j + 2).Value = cell.Value
                    current_sh.Cells(i, j + 3).Value = source_sh.Cells(cell.Row, col_ind + 1).Value
                    i = i + 1
                End If

            Next cell
            j = j + 5   'Add column break between different columns
            i = 2       'Reset row position for new block of data
        End If
    Next col_ind
    
Cleanup:
    Application.ScreenUpdating = True
    If Not source_wb Is Nothing Then source_wb.Close False
    MsgBox "Data transfer complete."
    Exit Sub

End Sub


