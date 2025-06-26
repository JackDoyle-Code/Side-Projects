Attribute VB_Name = "Module3"
' This sub is to be used when you want to count/sum/average a range of cells repeatedly but there is gaps between the next event of values (i.e. the data samples have a gap between each week)
' This sub will open up a separate file and sheet, take the first range of values (defined by count_cell_ref and rng_len/rng_wdt) and countif them. Then return the value into the active cell of your current file
' This sub can then be repeated each week and it will take into account the offset value of difference between each week

Sub current_cell_cal()

Dim filepath As String
Dim filepath_sh_num As Integer
Dim count_cell_ref As String
Dim offset_cell_ref As Integer
Dim offset_val As Integer
Dim current_cell As Integer
Dim rng_len As Integer
Dim rng_wdt As Integer
Dim output As Variant

' the sub section is only specific to the file path, however, the function itself can be used for other files as well
filepath = "Your file path"
filepath_sh_num = 2
count_cell_ref = "Your input cell in which you want to start the count from"
offset_cell_ref = Range("Your cell to offset from").Column
current_cell = ActiveCell.Column
offset_val = 23 'Your offset value
rng_len = 21 'Your range length
rng_wdt = 0 'Your range width (>0 if 2D range)
if_con = Array("Stocked", "half-stocked", "needs to be stocked") 'Your specified if condition

output = dynamic_count(filepath, filepath_sh_num, count_cell_ref, offset_cell_ref, current_cell, offset_val, rng_len, rng_wdt, if_con)
For i = 0 To UBound(output)
    ActiveCell.Offset(i, 0).Value = output(i)
Next i

End Sub

Function dynamic_count(filepath As String, filepath_sh_num As Integer, count_cell_ref As String, offset_cell_ref As Integer, current_cell As Integer, offset_val As Integer, rng_len As Integer, rng_wdt As Integer, if_con As Variant)
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim count As Integer
    Dim count_arr As Variant

    ' Open the workbook invisibly
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' sends the function to be resolved if there is an error
    On Error GoTo ErrorHandler
    ' opens the workbook and sheet
    Set wb = Workbooks.Open(filepath, ReadOnly:=True)
    Set ws = wb.Sheets(filepath_sh_num)

    ' Uses cell reference of the data you want to count and counts the next 21 rows (range of rows and columns). This is offset by 23 (the offset_val) * the number of times this sub has been called (current_cell).
    Set rng = ws.Range(ws.Cells(Range(count_cell_ref).Row + (offset_val * (current_cell - offset_cell_ref) / 7), Range(count_cell_ref).Column), _
                       ws.Cells(Range(count_cell_ref).Row + (offset_val * (current_cell - offset_cell_ref) / 7) + rng_len, Range(count_cell_ref).Column + rng_wdt))

    ' Count the "stocked" entries within the range
    ReDim count_arr(0 To UBound(if_con))
    For i = 0 To UBound(if_con)
        count_arr(i) = Application.WorksheetFunction.CountIf(rng, if_con(i))
    Next i
    
    dynamic_count = count_arr

    ' Closes the workbook after making the changes
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Function

' if there is an error it goes here and is resolved
ErrorHandler:
    ' sets the activecell as -1 to inform you that there is an error
    dynamic_count = -1
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
End Function



