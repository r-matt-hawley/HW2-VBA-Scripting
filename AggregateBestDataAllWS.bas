Attribute VB_Name = "AggregateBestDataAllWS"
Sub AggregateBestDataAllWS()

    'Columns storing original data
    Const StockTicker As Integer = 1
    Const StockDate As Integer = 2
    Const StockOpen As Integer = 3
    Const StockHigh As Integer = 4
    Const StockLow As Integer = 5
    Const StockClose As Integer = 6
    Const StockVolume As Integer = 7
    
    'Columns storing the Aggregated data
    'Note: Column 8 is blank for a visual buffer.
    Const SummaryTicker As Integer = 9
    Const SummaryYearlyChange As Integer = 10
    Const SummaryPercentChange As Integer = 11
    Const SummaryVolume As Integer = 12
    
    'Rows storing superlative data
    Const GreatestIncreaseRow As Integer = 2
    Const GreatestDecreaseRow As Integer = 3
    Const GreatestVolumeRow As Integer = 4
    
    'Columns storing superlative data
    Const SuperlativeHeader As Integer = 15
    Const SuperlativeTicker As Integer = 16
    Const SuperlativeData As Integer = 17

    Dim ActiveTicker As String
    Dim NextTicker As String
    Dim TotalVolume As Double
    Dim OpenValue As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim SummaryRow As Double
    Dim DataRow As Double
    Dim LastRow As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    Dim Current As Worksheet
    
    '---------------------------
    'Loop through all worksheets
    '---------------------------
    For Each Current In Worksheets
    
        'Find the last row of the data.
        LastRow = Current.Range("A" & Rows.Count).End(xlUp).Row
        
        'Set SummaryRow as the first row in the summary table.
        SummaryRow = 2
        
        'Set up Summary Table.
        Current.Cells(1, SummaryTicker).Value = "Ticker"
        Current.Cells(1, SummaryYearlyChange).Value = "Yearly Change"
        Current.Cells(1, SummaryPercentChange).Value = "Percent Change"
        Current.Cells(1, SummaryVolume).Value = "Total Stock Volume"
        
        'Set the opening share value for the first company.
        OpenValue = Current.Cells(2, StockOpen).Value
        
        'Initiate the superlative data.
        GreatestIncrease = 0
        LeastIncrease = 0
        GreatestVolume = 0
        
        '-----------------------
        'Loop through stock data
        '-----------------------
        For DataRow = 2 To LastRow
            
            'Always add to TotalVolume.
            TotalVolume = TotalVolume + Current.Cells(DataRow, StockVolume)
            
            'If OpenValue is zero update the opening share value.
            If (OpenValue = 0) Then OpenValue = Current.Cells(DataRow, StockOpen).Value
            
            '--------------
            'End of Company
            '--------------
            
            'If the next Ticker is different, then we have aggregated all of the data
            'for a single company.
            If (Current.Cells(DataRow, StockTicker).Value <> Current.Cells(DataRow + 1, StockTicker).Value) Then
                
                '-------------------------------------------------
                'Check if there is data available for the company.
                '-------------------------------------------------
                If (OpenValue = 0) Then
                
                    'Record the available data in the summary table, but
                    'note that the data is unavailable.
                    
                    'Ticker
                    Current.Cells(SummaryRow, SummaryTicker).Value = Current.Cells(DataRow, StockTicker).Value
                    
                    'YearlyChange
                    YearlyChange = Current.Cells(DataRow, StockClose).Value - OpenValue
                    Current.Cells(SummaryRow, SummaryYearlyChange).Value = YearlyChange
                    
                    'PercentChange: 0.00%
                    Current.Cells(SummaryRow, SummaryPercentChange).Value = "0.00%"
                    
                    'Yearly Volume
                    Current.Cells(SummaryRow, SummaryVolume).Value = TotalVolume
                    
                    'Color the cells with missing data gray signifying missing data.
                    For i = SummaryTicker To SummaryVolume
                    
                        Current.Cells(SummaryRow, i).Interior.ColorIndex = 48
                    
                    Next i
                    
                    'Place the error note in the column after Total Volume.
                    Current.Cells(SummaryRow, SummaryVolume + 1).Value = "Note: Data unavailable"
                    
                    'Resize the error note column.
                    Current.Columns(SummaryVolume + 1).AutoFit
                    
                     'Next time use the next summary row.
                    SummaryRow = SummaryRow + 1
                    
                    'Set the opening share value for the next company.
                    OpenValue = Current.Cells(DataRow + 1, StockOpen).Value
                    
                    'Reset the total volume.
                    TotalVolume = 0
                    
                
                Else
                    '-------------------------------------
                    'Record the data in the summary table.
                    '-------------------------------------
                    
                    'Ticker
                    Current.Cells(SummaryRow, SummaryTicker).Value = Current.Cells(DataRow, StockTicker).Value
                    
                    'Yearly Change: Close - Open
                    YearlyChange = Current.Cells(DataRow, StockClose).Value - OpenValue
                    Current.Cells(SummaryRow, SummaryYearlyChange).Value = YearlyChange
                    
                    'Color the yearly change Green if its value increased or
                    'red if its value decreased.
                    If (YearlyChange > 0) Then
                        Current.Cells(SummaryRow, SummaryYearlyChange).Interior.ColorIndex = 4 'Green
                    Else
                        Current.Cells(SummaryRow, SummaryYearlyChange).Interior.ColorIndex = 3 'Red
                    End If
                    
                    'Percent Change: (Close - Open) / Open
                    PercentChange = YearlyChange / OpenValue
                    Current.Cells(SummaryRow, SummaryPercentChange).Value = FormatPercent(PercentChange, 2)
                    
                    'Yearly Volume
                    Current.Cells(SummaryRow, SummaryVolume).Value = TotalVolume
                    
                    '----------------
                    'Superlative Data
                    '----------------
                    
                    'Check for Greatest Increase.
                    If (PercentChange > GreatestIncrease) Then
                        GreatestIncrease = PercentChange
                        GreatestIncreaseTicker = Current.Cells(DataRow, StockTicker).Value
                    End If
                    
                    'Check for Greatest Decrease.
                    If (PercentChange < GreatestDecrease) Then
                        GreatestDecrease = PercentChange
                        GreatestDecreaseTicker = Current.Cells(DataRow, StockTicker).Value
                    End If
                    
                    'Check for Greatest Volume.
                    If (TotalVolume > GreatestVolume) Then
                        GreatestVolume = TotalVolume
                        GreatestVolumeTicker = Current.Cells(DataRow, StockTicker).Value
                    End If
                    
                    '--------------
                    'Next Row Setup
                    '--------------
                    
                    'Next time, use the next summary row.
                    SummaryRow = SummaryRow + 1
                    
                    'Set the opening share value for the next company.
                    OpenValue = Current.Cells(DataRow + 1, StockOpen).Value
                    
                    'Reset the total volume
                    TotalVolume = 0
    
                End If
                
            '----------------
            'Summary Complete
            '----------------
            End If
            
        Next DataRow
         
        'Autofit each column in the summary table.
        For i = SummaryTicker To SummaryVolume
            Current.Columns(i).AutoFit
        Next i
        
        '------------------------
        'Record superlative data.
        '------------------------
        Current.Cells(1, SuperlativeTicker).Value = "Ticker"
        Current.Cells(1, SuperlativeData).Value = "Value"
        
        Current.Cells(GreatestIncreaseRow, SuperlativeHeader).Value = "Greatest % Increase"
        Current.Cells(GreatestIncreaseRow, SuperlativeTicker).Value = GreatestIncreaseTicker
        Current.Cells(GreatestIncreaseRow, SuperlativeData).Value = FormatPercent(GreatestIncrease, 2)
        
        Current.Cells(GreatestDecreaseRow, SuperlativeHeader).Value = "Greatest % Decrease"
        Current.Cells(GreatestDecreaseRow, SuperlativeTicker).Value = GreatestDecreaseTicker
        Current.Cells(GreatestDecreaseRow, SuperlativeData).Value = FormatPercent(GreatestDecrease, 2)
        
        Current.Cells(GreatestVolumeRow, SuperlativeHeader).Value = "Greatest Total Volume"
        Current.Cells(GreatestVolumeRow, SuperlativeTicker).Value = GreatestVolumeTicker
        Current.Cells(GreatestVolumeRow, SuperlativeData).Value = GreatestVolume
        
        'Clear superlative data for next worksheet.
        GreatestIncreaseTicker = ""
        GreatestIncrease = 0
        
        GreatestDecreaseTicker = ""
        GreatestDecrease = 0
        
        GreatestVolumeTicker = ""
        GreatestVolume = 0
        
        'Autofit each column in the superlative table.
        For i = SuperlativeHeader To SuperlativeData
            Current.Columns(i).AutoFit
        Next i
        
    '--------------
    'Next Worksheet
    '--------------
    Next Current
    
End Sub


