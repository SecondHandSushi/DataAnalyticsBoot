Sub challenge2()

Dim curr_tckr As String 'current ticker
Dim year_vol_tot As LongLong 'total volume for the year
Dim year_open As Variant 'opening value for the year
Dim year_close As Variant 'closing valuefor the year
Dim year_chng As Variant 'yearly change
Dim year_chng_perc As Variant 'yearly change as a percent
Dim col_tckr As Integer 'ticker column
Dim col_open As Integer 'opening column
Dim col_close As Integer 'closing col
Dim grt_incr_tckr As String 'greatest increase ticker
Dim grt_incr As Variant 'greatest increase value
Dim grt_decr_tckr As String 'greatest decrease ticker
Dim grt_decr As Variant 'greatest decrease value
Dim grt_vol_tot_tckr As String 'greatest total volume ticker
Dim grt_vol_tot As LongLong 'greatest total volume value
Dim eval_started As Boolean 'evaluation started
Dim ws As Worksheet

For Each ws In Worksheets 'for each statement that runs this or each worksheet in the workbook
    ws.Activate
    
'Initialize variables
grt_vol = 0
grt_incr = 0
grt_decr = 0
year_chng_perc = 0
year_chng = 0
year_open = 0
year_close = 0
year_vol_tot = 0
eval_started = 0


col_tckr = 1
col_open = 3
col_close = 6
col_vol = 7

i = 2    'this is the row index value for data retrieval
j = 1    'this is the row  index value to display results


'create colume headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Biggest % Gain"
Cells(3, 15).Value = "Biggest % Loss"
Cells(4, 15).Value = "Greatest Total Volume"


Do 'loop is active until an empty cell is detected
    If eval_started = False Then 'new ticker detected
        year_open = Cells(i, col_open).Value
        year_vol_tot = year_vol_tot + Cells(i, col_vol).Value
        eval_started = True
        i = i + 1
        
    ElseIf Cells(i + 1, col_tckr).Value <> Cells(i, col_tckr).Value Then 'current and next cell are different, nearly all data gathering is considered complete, now it is evaluated
        j = j + 1
        year_vol_tot = year_vol_tot + Cells(i, col_vol).Value 'add last volume value to yearly total
        year_close = Cells(i, col_close).Value
        year_chng = year_close - year_open
        year_chng_perc = (year_close / year_open) - 1
        curr_tckr = Cells(i, col_tckr).Value
        

        'compare and replace greatest increase with current value if greater
        'if not then the greatest increase value remains unchanged
        If year_chng_perc > grt_incr Then
                grt_incr = year_chng_perc
                grt_incr_tckr = curr_tckr
        ElseIf year_chng_perc < grt_incr Then
                grt_incr = grt_incr
        End If
        
        'compare and replace greatest decrease with current value if it is less
        'if not then the greatest decrease value remains unchanged
        If year_chng_perc < grt_decr Then
                grt_decr = year_chng_perc
                grt_decr_tckr = curr_tckr
        ElseIf year_chng_perc > grt_decr Then
                grt_decr = grt_decr
        End If
        
        'compare and replace greatest volume with current value if it is greater
        'if not then the greatest decrease value remains unchanged
        If year_vol_tot > grt_vol Then
                grt_vol = year_vol_tot
                grt_vol_tot_tckr = curr_tckr
        ElseIf year_vol_tot < grt_vol Then
                grt_vol = grt_vol
        End If
        
        'display yearly results
        Cells(j, 9).Value = curr_tckr
        Cells(j, 10).Value = year_chng
        Cells(j, 11).Value = year_chng_perc
        Cells(j, 12).Value = year_vol_tot
        Cells(2, 16).Value = grt_incr_tckr
        Cells(3, 16).Value = grt_decr_tckr
        Cells(4, 16).Value = grt_vol_tot_tckr
        
        'cell formatting
        If (year_chng > 0) Then           'format cell color to match positive or negative yearly change
                Cells(j, 10).Interior.ColorIndex = 4
        ElseIf (year_chng < 0) Then
                Cells(j, 10).Interior.ColorIndex = 3
        End If
        Cells(j, 11).NumberFormat = "0.00%" 'format cells to percent with two decimal points
        Range("Q2:Q3").NumberFormat = "0.00%" '
        
        'display greatest value results
        Cells(2, 17).Value = grt_incr
        Cells(3, 17).Value = grt_decr
        Cells(4, 17).Value = grt_vol
        
        'reinitialize variables for next loop
        year_vol_tot = 0
        i = i + 1
        eval_started = False
    Else 'no new ticker detected
        year_vol_tot = year_vol_tot + Cells(i, col_vol).Value
        i = i + 1
    End If
Loop While Cells(i, col_tckr).Value <> ""

Next    'run script on next worksheet
End Sub
