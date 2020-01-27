Attribute VB_Name = "Multiyear_Easy"
Sub Multiyear_Easy()

'To find out start time
Dim startTime As Date
Dim endTime As Date
startTime = Time()
Debug.Print "start time is " & startTime

    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    'To print worksheet name
    ws_name = ws.Name
    Debug.Print "worksheet name is " & ws_name
    
        Dim sourceCol As Long
        Dim destCol As Long
        Dim distinct_ticker As Range
        Dim all_ticker As Range
        Dim row_count As Long
        Dim header_row_count As Long
        Dim total_volume As Variant
        Dim volCol As Long
        Dim column_volume As Long
        Dim result_row_count As Long
        
        sourceCol = 1           'Source Ticker column
        destCol = 9             'Destination Ticker Column
        volCol = 7              'Source Volume Column
        header_row_count = 1    'To print result headers
        row_count = 2           'Data row count
        result_row_count = 2

        'print header "Ticker" and header "Total Stock Volume"
        ws.Cells(header_row_count, destCol) = "Ticker"
        ws.Cells(header_row_count, destCol + 1) = "Total Stock Volume"
       
        total_volume = 0
           
        'Inner loop to loop through all tickers from source column
        'Inner loop will run till last row of all_ticker column to search for out of position tickers
        For Each all_ticker In ws.Range(ws.Cells(row_count, sourceCol), ws.Cells(Rows.Count, sourceCol).End(xlUp)).Cells
        
                column_volume = CLng(ws.Cells(all_ticker.Row, volCol).Value)
                'Debug.Print "column_volume is " & column_volume & "this"
                total_volume = total_volume + column_volume
                
                'Debug.Print "all_ticker is " & all_ticker.Value
                'Debug.Print "next ticker is " & ws.Cells(all_ticker.Row + 1, sourceCol).Value
                
                If ws.Cells(all_ticker.Row + 1, sourceCol).Value <> all_ticker.Value Then
                ws.Cells(result_row_count, destCol) = all_ticker.Value
                'Debug.Print "all_ticker is " & all_ticker.Value
                'Debug.Print "total_volume is " & total_volume
                ws.Cells(result_row_count, destCol + 1) = total_volume
                total_volume = 0
                result_row_count = result_row_count + 1
                End If
                                                   
        Next all_ticker
        
   'Auto fit all columns
   ws.Cells.EntireColumn.AutoFit
        
'To find out end time
endTime = Time()
Debug.Print "end time is " & endTime
End Sub



