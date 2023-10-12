Sub challenge2()

Dim current_ticker As String
Dim year_volume_total As LongLong
Dim year_open As Variant
Dim year_close As Variant
Dim year_change As Variant
Dim year_change_percent As Variant
Dim column_ticker As Integer
Dim column_open As Integer
Dim column_close As Integer
Dim greatest_increase_ticker As String
Dim greatest_increase As Variant
Dim greatest_decrease_ticker As String
Dim greatest_decrease As Variant
Dim greatest_volume_ticker As String
Dim greatest_volume As LongLong
Dim sample_count As Long
Dim evaluation_started As Boolean
Dim ws As Worksheet

For Each ws In Worksheets 'for each statement that runs this or each worksheet in the workbook
    ws.Activate
    
'Initialize variables
greatest_volume = 0
greatest_increase = 0
greatest_decrease = 0
year_change_percent = 0
year_change = 0
year_open = 0
year_close = 0
year_volume_total = 0
evluation_started = 0


column_ticker = 1
column_open = 3
column_close = 6
column_volume = 7

i = 2    'this is the row index value for data retrieval
j = 1    'this is the row  index value to display results


'create colume headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Biggest Gain"
Cells(3, 15).Value = "Biggest Loss"
Cells(4, 15).Value = "Largest Volume"


Do 'loop is active until an empty cell is detected
    If evaluation_started = False Then 'new ticker detected
        year_open = Cells(i, column_open).Value
        year_volume_total = year_volume_total + Cells(i, column_volume).Value
        evaluation_started = True
        i = i + 1
        
    ElseIf Cells(i + 1, column_ticker).Value <> Cells(i, column_ticker).Value Then 'current and next cell are different, nearly all data gathering is considered complete, now it is evaluated
        j = j + 1
        year_volume_total = year_volume_total + Cells(i, column_volume).Value 'add last volume value to yearly total
        year_close = Cells(i, column_close).Value
        year_change = year_close - year_open
        year_change_percent = (year_close / year_open) - 1
        current_ticker = Cells(i, column_ticker).Value
        

        'compare and replace greatest increase with current value if greater
        'if not then the greatest increase value remains unchanged
        If year_change_percent > greatest_increase Then
                greatest_increase = year_change_percent
                greatest_increase_ticker = current_ticker
        ElseIf year_change_percent < greatest_increase Then
                greatest_increase = greatest_increase
        End If
        
        'compare and replace greatest decrease with current value if it is less
        'if not then the greatest decrease value remains unchanged
        If year_change_percent < greatest_decrease Then
                greatest_decrease = year_change_percent
                greatest_decrease_ticker = current_ticker
        ElseIf year_change_percent > greatest_decrease Then
                greatest_decrease = greatest_decrease
        End If
        
        'compare and replace greatest volume with current value if it is greater
        'if not then the greatest decrease value remains unchanged
        If year_volume_total > greatest_volume Then
                greatest_volume = year_volume_total
                greatest_volume_ticker = current_ticker
        ElseIf year_volume_total < greatest_volume Then
                greatest_volume = greatest_volume
        End If
        
        'display yearly results
        Cells(j, 9).Value = current_ticker
        Cells(j, 10).Value = year_change
        Cells(j, 11).Value = year_change_percent
        Cells(j, 12).Value = year_volume_total
        Cells(2, 16).Value = greatest_increase_ticker
        Cells(3, 16).Value = greatest_decrease_ticker
        Cells(4, 16).Value = greatest_volume_ticker
        
        'cell formatting
        If (year_change > 0) Then           'format cell color to match positive or negative yearly change
                Cells(j, 10).Interior.ColorIndex = 4
        ElseIf (year_change < 0) Then
                Cells(j, 10).Interior.ColorIndex = 3
        End If
        Cells(j, 11).NumberFormat = "0.00%" 'format cells to percent with two decimal points
        Range("Q2:Q3").NumberFormat = "0.00%" '
        
        'display greatest value results
        Cells(2, 17).Value = greatest_increase
        Cells(3, 17).Value = greatest_decrease
        Cells(4, 17).Value = greatest_volume
        
        'reinitialize variables for next loop
        year_volume_total = 0
        i = i + 1
        evaluation_started = False
    Else 'no new ticker detected
        year_volume_total = year_volume_total + Cells(i, column_volume).Value
        i = i + 1
    End If
Loop While Cells(i, column_ticker).Value <> ""

Next    'run script on next worksheet
End Sub
