Sub ticker()
    'creating all neccessary variables
    Dim start As Double
    Dim finish As Double
    Dim vol As Variant
    Dim lastrow As Variant
    Dim lastrowtable As Integer
    'creating the position variable for the table
    Dim pos As Integer
    'using the count variable as a way to get the start and end of a stock
    Dim count As Integer
    Dim great_increase As Double
    Dim great_decrease As Double
    Dim great_volume As Double
    Dim great_increase_name As String
    Dim great_decrease_name As String
    Dim great_volume_name As String
    
    For Each ws In Worksheets
        pos = 2
        'creating the table labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'getting the amount of data entries in dataset
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
        'creating the table with ticker, yearly change, percent change, and total stock volume
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                'adding to the total volume of the stock
                vol = vol + ws.Cells(i, 7).Value
                count = count + 1
            Else
                'adding the last volume of the stock
                vol = vol + ws.Cells(i, 7).Value
                'getting the first open and last close of the stock
                start = ws.Cells(i - count, 3).Value
                finish = ws.Cells(i, 6).Value
                'assigning the values to the outcome table
                ws.Cells(pos, 9).Value = Cells(i, 1).Value
                ws.Cells(pos, 10).Value = finish - start
                ws.Cells(pos, 11).Value = (finish - start) / start
                ws.Cells(pos, 12).Value = vol
                'resetting the vol and count
                vol = 0
                count = 0
                'going to the next row of the table
                pos = pos + 1
            End If
        Next i

        'changing format of Percent Change and Yearly Change columns
        ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
        ws.Range("J2:J" & lastrow).NumberFormat = "0.00"

        'formatting Yearly Change column
        lastrowtable = ws.Cells(Rows.count, 10).End(xlUp).Row
    
        For i = 2 To lastrowtable
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
        Next i
        
        'finding the Greatest % Increase, Greatest % decrease and, Greatest total Volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
                
        great_increase = ws.Cells(2, 11).Value
        greart_decrease = ws.Cells(2, 11).Value
        great_volume = ws.Cells(2, 12).Value
        
        For i = 2 To lastrowtable
            If ws.Cells(i, 11).Value > great_increase Then
                great_increase = ws.Cells(i, 11).Value
                great_increase_name = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value < great_decrease Then
                great_decrease = ws.Cells(i, 11).Value
                great_decrease_name = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value > great_volume Then
                great_volume = ws.Cells(i, 12).Value
                great_volume_name = ws.Cells(i, 9).Value
            End If
        Next i
            
        'putting values into the table
        ws.Cells(2, 16).Value = great_increase_name
        ws.Cells(2, 17).Value = great_increase
        ws.Cells(3, 16).Value = great_decrease_name
        ws.Cells(3, 17).Value = great_decrease
        ws.Cells(4, 16).Value = great_volume_name
        ws.Cells(4, 17).Value = great_volume
    Next ws

End Sub
