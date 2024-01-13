Attribute VB_Name = "Module1"
Sub StockData()

    ' Loop through all Spreadsheets
    For Each ws In Worksheets

        ' Set an Initial Variables
        Dim TickerName As String
        Dim GreatestPercentIncreaseTickerName As String
        Dim GreatestPercentDecreaseTickerName As String
        Dim GreatestTotalVolumeTickerName As String
        Dim LastRow As Long
        Dim FirstDayOpeningPrice As Double
        Dim LastDayClosingingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As LongLong
        Dim GreatestPercentIncreaseTotal As Double
        Dim GreatestPercentDecreaseTotal As Double
        Dim GreatestTotalVolumeTotal As LongLong
        
        ' Initialize Variables
        TickerName = ""
        GreatestPercentIncreaseTickerName = ""
        GreatestPercentDecreaseTickerName = ""
        GreatestTotalVolumeTickerName = ""
        LastRow = 0
        FirstDayOpeningPrice = 0
        LastDayClosingingPrice = 0
        YearlyChange = 0
        PercentChange = 0
        TotalStockVolume = 0
        GreatestPercentIncreaseTotal = 0
        GreatestPercentDecreaseTotal = 0
        GreatestTotalVolumeTotal = 0
        
           
        ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Integer
        
        ' Initialize Summary_Table_Row to start on Row 2
        Summary_Table_Row = 2
        
        ' Set Sumary Table Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Set Greatest Table Headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        
        'Find LastRow on Sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
              
        ' Loop through all Tickers
        For i = 2 To LastRow
                   
            ' Check if we have changed TickerName
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the TickerName
                TickerName = ws.Cells(i, 1).Value
    
                ' Set LastDayClosingingPrice
                LastDayClosingingPrice = ws.Cells(i, 6).Value
    
                ' Add to the TotalStockVolume
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                ' Calculate YearlyChange
                YearlyChange = LastDayClosingingPrice - FirstDayOpeningPrice
                
                ' Calculate PercentChange
                PercentChange = YearlyChange / FirstDayOpeningPrice
                
                                             
                                             
                ' Check to see if values are greatest on the sheet so far, if so update
                '----------------------------------------------------------------------
                If PercentChange > GreatestPercentIncreaseTotal Then
                    GreatestPercentIncreaseTotal = PercentChange
                    GreatestPercentIncreaseTickerName = TickerName
                End If
                
                If PercentChange < GreatestPercentDecreaseTotal Then
                    GreatestPercentDecreaseTotal = PercentChange
                    GreatestPercentDecreaseTickerName = TickerName
                End If
                
                If TotalStockVolume > GreatestTotalVolumeTotal Then
                    GreatestTotalVolumeTotal = TotalStockVolume
                    GreatestTotalVolumeTickerName = TickerName
                End If
                
                
                    
                ' Inputting data into Summary Table
                ' ---------------------------------
                ' Print the TickerName in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = TickerName
                
                ' Print the YearlyChange in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                
                ' Set YearlyChange Color
                If YearlyChange >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                End If
                
                ' Print the PercentChange in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                
                ' Format PercentChange to a %
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ' Print the TotalStockVolume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
                
                                    
                                    
                ' Reset for next Ticker
                '---------------------------
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
          
                ' Reset TotalStockVolume
                TotalStockVolume = 0
                
                ' Reset FirstDayOpeningPrice
                FirstDayOpeningPrice = 0
                
                ' Reset YearlyChange
                YearlyChange = 0
                
                ' Reset PercentChange
                PercentChange = 0
    
            ' If next cell is the same TickerName
            Else
                If TotalStockVolume = 0 Then
                    FirstDayOpeningPrice = ws.Cells(i, 3).Value
                End If
                
                ' Add to the TotalStockVolume
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                           
            End If
    
        Next i
        
        
        ' Print out Greatest Ticker Names and Totals for the sheet
        ' --------------------------------------------------------
        ws.Range("P2").Value = GreatestPercentIncreaseTickerName
        ws.Range("Q2").Value = GreatestPercentIncreaseTotal
        
        ws.Range("P3").Value = GreatestPercentDecreaseTickerName
        ws.Range("Q3").Value = GreatestPercentDecreaseTotal
        
        ws.Range("P4").Value = GreatestTotalVolumeTickerName
        ws.Range("Q4").Value = GreatestTotalVolumeTotal
        
        ' Format GreatestPercents to a %
        ws.Range("Q2,Q3").NumberFormat = "0.00%"
               
        ' Set all columns to automatically adjust to proper width
        ws.Cells.EntireColumn.AutoFit
        
    Next ws

End Sub

