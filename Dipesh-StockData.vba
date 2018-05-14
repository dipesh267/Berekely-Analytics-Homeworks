Sub TickerTotal():
    Dim total As Double
    Dim opener As Double
    Dim closer As Double
    Dim change As Double
    Dim percent_change As Double
    
    Dim greatest_increase_ticker As String
    Dim greatest_increase As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    
    Dim ticker_name As String
    
    Dim counter As Double
    counter = 2
    
    Dim lastrow As Double
    Dim lastrow_mini As Integer
    
    Dim WS_Count As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For Each ws In Worksheets
        ws.Activate
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            If i = 2 Then
                    'grab the opening amount
                    opener = Cells(i, 3).Value
            End If
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker_name = Cells(i, 1).Value
                total = total + Cells(i, 7).Value
                closer = Cells(i, 6).Value
                change = closer - opener
                
                If opener <> 0 Then
                    percentage_change = change / opener
                End If
                
                Cells(counter, 9).Value = ticker_name
                Cells(counter, 12).Value = total
                'change opener for next Ticker
                opener = Cells(i + 1, 3).Value
                
                Cells(counter, 10).Value = change
                If Cells(counter, 10).Value >= 0 Then
                    Cells(counter, 10).Interior.ColorIndex = 4
                Else
                    Cells(counter, 10).Interior.ColorIndex = 3
                End If
                
                Cells(counter, 11).Value = percentage_change
                Cells(counter, 11).NumberFormat = "0.00%"
                counter = counter + 1
                total = 0
            Else
                total = total + Cells(i, 7).Value
            End If
        Next i
        counter = 2
        i = 2
        
        'look for the greatest values from the mini table
        'grab the row count of mini table
        lastrow_mini = Cells(Rows.Count, 9).End(xlUp).Row
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0
        Dim temp As Double
        Dim temp_vol As Double
        For j = 2 To lastrow_mini
            temp = Cells(j, 11).Value
            temp_vol = Cells(j, 12).Value
            'grab the biggest percentage
            If (temp > greatest_increase) Then
                greatest_increase = temp
                greatest_increase_ticker = Cells(j, 9).Value
            End If
            'grab the lowest percentage
            If (temp < greatest_decrease) Then
                greatest_decrease = temp
                greatest_decrease_ticker = Cells(j, 9).Value
            End If
            'grab the biggets volume
            If (temp_vol > greatest_volume) Then
                greatest_volume = temp_vol
                greatest_volume_ticker = Cells(j, 9).Value
            End If
            
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percentage Change"
            Range("L1").Value = "Total Volume"
            
            Range("N2").Value = "Greatest % Increase"
            Range("N3").Value = "Greatest % Decrease"
            Range("N4").Value = "Greatest Total Volume"
            
            Range("O1").Value = "Ticker"
            Range("O2").Value = greatest_increase_ticker
            Range("P1").Value = "Value"
            Range("P2").Value = greatest_increase
            Range("P2").NumberFormat = "0.00%"
            
            Range("O3").Value = greatest_decrease_ticker
            Range("P3").Value = greatest_decrease
            Range("P3").NumberFormat = "0.00%"
            Range("O4").Value = greatest_volume_ticker
            Range("P4").Value = greatest_volume
        Next j
    Next ws
    
End Sub

