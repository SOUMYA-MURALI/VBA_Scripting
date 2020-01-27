Attribute VB_Name = "Multiyear_Moderate_Challenge"
Sub Multiyear_Moderate_Challenge()

Dim startTime As Date
Dim endTime As Date
startTime = Time()
Debug.Print "start time is " & startTime

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws_name = ws.Name
    Debug.Print "worksheet name is " & ws_name
    
        Dim sourceCol As Long
        Dim destCol As Long
        Dim distinct_ticker As Range
        Dim all_ticker As Range
        Dim total_count As Long
        Dim row_count As Long
        Dim header_row_count As Long
        Dim total_volume As Variant
        Dim volCol As Long
        Dim column_volume As Long
        Dim column_date As Long
        Dim dateCol As Long
        Dim min_column_date As Long
        Dim max_column_date As Long
        Dim min_open_price As Double
        Dim max_close_price As Double
        Dim openCol As Long
        Dim closeCol As Long
        Dim percent_change As Double
        Dim yearly_change As Double
        Dim percent_change_format As String
        Dim yearly_rg As Range
        
        sourceCol = 1           'Source Ticker column
        destCol = 9             'Destination Ticker Column
        volCol = 7              'Source Volume Column
        dateCol = 2             'Source date column
        header_row_count = 1    'To print result headers
        row_count = 2           'Data row count
        openCol = 3             'Source open price column
        closeCol = 6            'Source close price column
        result_row_count = 2    'Result data row count
                 
        'Hard coding necessary headings
        ws.Cells(header_row_count, destCol) = "Ticker"
        ws.Cells(header_row_count, destCol + 1) = "Yearly Change"
        ws.Cells(header_row_count, destCol + 2) = "Percent Change"
        ws.Cells(header_row_count, destCol + 3) = "Total Stock Volume"

        total_volume = 0
        
        'Inner loop to loop through all tickers from source column
        'Inner loop will run till last row of all_ticker column to search for out of position tickers
        For Each all_ticker In ws.Range(ws.Cells(row_count, sourceCol), ws.Cells(Rows.Count, sourceCol).End(xlUp)).Cells
        
                column_volume = CLng(ws.Cells(all_ticker.Row, volCol).Value)
                column_date = CLng(ws.Cells(all_ticker.Row, dateCol).Value)
                'Debug.Print "column_volume is " & column_volume & "this"
                total_volume = total_volume + column_volume
                
                If min_column_date = 0 Then
                        min_column_date = column_date
                        min_open_price = Format(ws.Cells(all_ticker.Row, openCol).Value, "0000.000000000")
                ElseIf column_date < min_column_date Then
                        min_column_date = column_date
                        min_open_price = ws.Cells(all_ticker.Row, openCol).Value
                        min_open_price = Format(ws.Cells(all_ticker.Row, openCol).Value, "0000.000000000")
                End If
                
                If max_column_date = 0 Then
                        max_column_date = column_date
                        max_close_price = Format(ws.Cells(all_ticker.Row, closeCol).Value, "0000.000000000")
                ElseIf column_date > max_column_date Then
                        max_column_date = column_date
                        max_close_price = Format(ws.Cells(all_ticker.Row, closeCol).Value, "0000.000000000")
                        
                End If
                
                'Debug.Print "all_ticker is " & all_ticker.Value
                'Debug.Print "next ticker is " & ws.Cells(all_ticker.Row + 1, sourceCol).Value
                
                If ws.Cells(all_ticker.Row + 1, sourceCol).Value <> all_ticker.Value Then
                
                ws.Cells(result_row_count, destCol) = all_ticker.Value
                'Debug.Print "all_ticker is " & all_ticker.Value
                'Debug.Print "total_volume is " & total_volume
                ws.Cells(result_row_count, destCol + 3) = total_volume
                            
                yearly_change = Format((max_close_price - min_open_price), "0000.000000000")
                ws.Cells(result_row_count, destCol + 1) = Format(yearly_change, "0000.000000000")
                
                If min_open_price = 0 Then
                percent_change = 0
                Else
                'percent_change = Format(((max_close_price - min_open_price) / (min_open_price)) * 100, "####.##")
                percent_change = ((max_close_price - min_open_price) / (min_open_price))
                End If
                'percent_change_String = percent_change & "%"
                
                percent_change_format = Format(percent_change, "000.00%")
                            
                ws.Cells(result_row_count, destCol + 2) = percent_change_format
            
                total_volume = 0
                result_row_count = result_row_count + 1
                min_column_date = 0
                max_column_date = 0
                                        
                End If
                                                   
        Next all_ticker
        
        Dim cond_1 As FormatCondition, cond_2 As FormatCondition, cond_3 As FormatCondition
        Set rg = ws.Range(ws.Cells(2, destCol + 1), ws.Cells(Rows.Count, destCol + 1).End(xlUp)).Cells
        
        Set cond_1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "0")
        Set cond_2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")
        Set cond_3 = rg.FormatConditions.Add(xlCellValue, xlEqual, "0")
        
        With cond_1
        .Interior.Color = vbGreen
        End With
        
        With cond_2
        .Interior.Color = vbRed
        End With
        
        With cond_3
        .Interior.Color = vbGreen
        End With
         
   'Auto fit all columns
   ws.Cells.EntireColumn.AutoFit
        
Next ws
endTime = Time()
Debug.Print "end time is " & endTime
End Sub





