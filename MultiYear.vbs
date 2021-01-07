Sub Tickers():
  Dim ws As Worksheet
      
  For Each ws In Worksheets
    'Create Columns
        
        Dim Symbol As String
        Symbol = "Ticker Symbol"
        ws.Range("I1").Value = Symbol
    
        Dim YearChange As String
        YearChange = "Yearly Change"
        ws.Range("J1").Value = YearChange
    
        Dim Percent As String
        Percent = "Percent Change"
        ws.Range("K1").Value = Percent
    
        Dim Stock As String
        Stock = "Total Stock Volume"
        ws.Range("L1").Value = Stock
    
        Dim OpeningPrice As String
        OpeningPrice = "Opening Price"
        ws.Range("N1").Value = OpeningPrice
    
        Dim ClosingPrice As String
        ClosingPrice = "Closing Price"
        ws.Range("O1").Value = ClosingPrice
        
    'Find Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Pull First Opening Price
        Dim OpeningPrice_Counter As Integer
        OpeningPrice_Counter = 2
        Dim OpenPrice As Double
        OpenPrice = ws.Cells(2, 3).Value
        ws.Range("N" & OpeningPrice_Counter).Value = OpenPrice
        OpeningPrice_Counter = OpeningPrice_Counter + 1
        
    'Pull Remaining Opening Prices
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            OpenPrice = ws.Cells(i + 1, 3).Value
            ws.Range("N" & OpeningPrice_Counter).Value = OpenPrice
            OpeningPrice_Counter = OpeningPrice_Counter + 1
            
        End If
        Next i
        
    'Pull Closing Price
        Dim ClosingPrice_Counter As Integer
        ClosingPrice_Counter = 2
        Dim ClosePrice As Double
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ClosePrice = ws.Cells(i, 6).Value
            ws.Range("O" & ClosingPrice_Counter).Value = ClosePrice
            ClosingPrice_Counter = ClosingPrice_Counter + 1
        End If
        Next i
    
    'Pull Ticker Symbol and Total Stock Volume
        Dim TickerSymbol As String
        Dim TickerCount As Integer
        TickerCount = 2
        Dim StockVol As LongLong
        StockVol = 0
    
            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerSymbol = ws.Cells(i, 1).Value
                ws.Range("I" & TickerCount).Value = TickerSymbol
                StockVol = StockVol + ws.Cells(i, 7).Value
                ws.Range("L" & TickerCount).Value = StockVol
                TickerCount = TickerCount + 1
        
                StockVol = 0
        
                ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                StockVol = StockVol + ws.Cells(i, 7).Value
        
                End If
        
            Next i
            
    'Find LastTicker
        LastTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
    'Pull Yearly Change
        TickerCount = 2
        OpeningPrice_Counter = 2
        ClosingPrice_Counter = 2
        
        Dim YearlyChange As Double
         
        For j = 2 To LastTicker
            ClosePrice = ws.Range("O" & ClosingPrice_Counter).Value
            OpenPrice = ws.Range("N" & OpeningPrice_Counter).Value
        
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J" & TickerCount).Value = YearlyChange
                If YearlyChange > 0 Then
                ws.Range("J" & TickerCount).Interior.ColorIndex = 4
                    
                ElseIf YearlyChange < 0 Then
                ws.Range("J" & TickerCount).Interior.ColorIndex = 3
                    
                End If
            TickerCount = TickerCount + 1
            OpeningPrice_Counter = OpeningPrice_Counter + 1
            ClosingPrice_Counter = ClosingPrice_Counter + 1
        
        Next j
        
    'Pull Percent Change
        TickerCount = 2
        OpeningPrice_Counter = 2
        ClosingPrice_Counter = 2
        
        Dim PercentChange As Double
        Dim PercentDiff As Double
        Dim DivPrice As Double
                
        For j = 2 To LastTicker
            ClosePrice = ws.Range("O" & ClosingPrice_Counter).Value
            OpenPrice = ws.Range("N" & OpeningPrice_Counter).Value
            DivPrice = ws.Range("N" & OpeningPrice_Counter).Value
            
            If (ClosePrice = 0) And (OpenPrice = 0) And (DivPrice = 0) Then
            PercentChange = 0
            
            ElseIf (ClosePrice <> 0) And (OpenPrice <> 0) And (DivPrice <> 0) Then
            PercentChange = (ClosePrice - OpenPrice) / DivPrice
            End If
            
            ws.Range("K" & TickerCount).Value = PercentChange
                
                If PercentChange > 0 Then
                ws.Range("K" & TickerCount).Interior.ColorIndex = 4
                    
                ElseIf PercentChange < 0 Then
                ws.Range("K" & TickerCount).Interior.ColorIndex = 3
                    
                End If
            
            TickerCount = TickerCount + 1
            OpeningPrice_Counter = OpeningPrice_Counter + 1
            ClosingPrice_Counter = ClosingPrice_Counter + 1
              
        Next j
            
     Next ws
End Sub
