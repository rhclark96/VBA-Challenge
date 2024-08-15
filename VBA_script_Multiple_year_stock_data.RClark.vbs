Sub stock_analysis()

    'Declare variables
    Dim i As Double
    Dim r As Double
    Dim table_row As Long
    Dim stock_volume_total As Double
    Dim ws As Worksheet
    Dim last_row As Double
    Dim date_value As Double
    Dim converted_date As Date
    Dim ticker As String
    Dim closing_price As Double
    Dim opening_price As Double
    Dim opening_price_stored As Boolean
    Dim percent_change As Double
    Dim summary_table_last_row As Double
    Dim max_percent_increase As Double
    Dim max_percent_decrease As Double
    Dim max_total_volume As Double

    'Set initial values
    table_row = 2
    stock_volume_total = 0

    'Set initial value of opening price stored boolean
    opening_price_stored = False

    ' Loop through worksheets
    For Each ws In ThisWorkbook.Worksheets

        'Set last row
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        'Label columns for table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Label max/min table rows and columns
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "ticker"
        ws.Cells(1, 16).Value = "value"

        ' Loop through tickers
        For i = 2 To last_row

            ' Convert dates (if necessary)
            'date_value = ws.Cells(i, 2).Value
            'converted_date = DateSerial(Left(date_value, 4), Mid(date_value, 5, 2), Right(date_value, 2))
            'ws.Cells(i, 2).Value = converted_date

            ' Group by ticker name
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                
                'Add to stick volume total
                stock_volume_total = stock_volume_total + ws.Cells(i, 7).Value

                ' Store opening price using boolean (Inspired by https://stackoverflow.com/questions/59461571/how-do-i-keep-initial-value-in-a-for-loop)
                If Not opening_price_stored Then
                    opening_price = ws.Cells(i, 3).Value
                    opening_price_stored = True
                End If

            Else
                ' Set ticker name
                ticker = ws.Cells(i, 1).Value
                ws.Cells(table_row, 9).Value = ticker

                ' Add to the stock volume total
                stock_volume_total = stock_volume_total + ws.Cells(i, 7).Value
                ws.Cells(table_row, 12).Value = stock_volume_total

                ' Set closing price
                closing_price = ws.Cells(i, 6).Value

                ' Print quarter/annual difference
                ws.Cells(table_row, 10).Value = closing_price - opening_price

                ' Print percent change
                percent_change = (ws.Cells(table_row, 10).Value / opening_price)
                ws.Cells(table_row, 11).Value = percent_change
                ws.Cells(table_row, 11).NumberFormat = "0.00%"


                ' Move to the next row
                table_row = table_row + 1

                ' Reset stock volume
                stock_volume_total = 0
                
                'Reset opening price boolean to false
                opening_price_stored = False
                
             End If

        Next i
        
            'Determine summary table last row
            summary_table_last_row = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
                'MsgBox (summary_table_last_row) (helped with debug)
            
            'Find max values for increase, decrease percent and max total volume
            max_percent_increase = WorksheetFunction.Max(ws.Range("K2:K" & summary_table_last_row))
                'MsgBox (max_percent_increase)
            max_percent_decrease = WorksheetFunction.Min(ws.Range("K2:K" & summary_table_last_row))
                'MsgBox (max_percent_decrease)
            max_total_volume = WorksheetFunction.Max(ws.Range("L2:L" & summary_table_last_row))
                'MsgBox (max_total_volume)
            
            'Print min/max values to table
            ws.Cells(2, 16).Value = max_percent_increase
            ws.Cells(3, 16).Value = max_percent_decrease
            ws.Range("P2:P3").NumberFormat = "0.00%"
            ws.Cells(4, 16).Value = max_total_volume
            
            'Loop through tickers in table to find ticker value matching min/max values
                For r = 2 To summary_table_last_row
                    If ws.Cells(r, 11).Value = max_percent_increase Then
                        ws.Cells(2, 15).Value = (ws.Cells(r, 9).Value)
                        
                    ElseIf ws.Cells(r, 11).Value = max_percent_decrease Then
                        ws.Cells(3, 15).Value = (ws.Cells(r, 9).Value)
                        
                    ElseIf ws.Cells(r, 12).Value = max_total_volume Then
                        ws.Cells(4, 15).Value = (ws.Cells(r, 9).Value)
                    
                    
                    End If
                Next r
        'Reset table rows for next ws
        table_row = 2
    Next ws

End Sub

