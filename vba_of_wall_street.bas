Attribute VB_Name = "Module1"
Sub tickersummary():

'Loop through all sheets

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
    
'Declare Variables
        Dim TickerName As String
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        Dim LastRow As Double
        Dim OpenPriceRow As Double
        Dim ColorRed As Integer
        Dim ColorGreen As Integer


'Set colors for conditional formatting
 
        ColorRed = 3
        ColorGreen = 4
        
'Set Opening Price row
        OpenPriceRow = 2
        
'Set last row for loop
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set starting total stock volume
        TotalStockVolume = 0
        
        Dim SummaryTableRow
        SummaryTableRow = 2
        
'Insert Header for summary table rows
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"

'Loop Through all rows in sheet
        For i = 2 To LastRow
        
        'Set condition to determine a when there is a new ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            
        'Ticker Symbol
            TickerName = Cells(i, 1)
        'Opening price
            OpeningPrice = Cells(OpenPriceRow, 3)
        'Closing Price
            ClosingPrice = Cells(i, 6)
        'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
            YearlyChange = ClosingPrice - OpeningPrice
        'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            PercentChange = YearlyChange / OpeningPrice
        'The total stock volume of the stock.
            TotalStockVolume = Cells(i, 7).Value + TotalStockVolume
            
        'Insert values into summary table
            Range("J" & SummaryTableRow).Value = TickerName
            Range("M" & SummaryTableRow).Value = TotalStockVolume
            Range("K" & SummaryTableRow).Value = YearlyChange
            
        'Conditional formatting that will highlight positive change in green and negative change in red.
                If YearlyChange < 0 Then
                Range("K" & SummaryTableRow).Interior.ColorIndex = ColorRed
                
                ElseIf YearlyChange > 0 Then
                Range("K" & SummaryTableRow).Interior.ColorIndex = ColorGreen
                
                End If
        'Format Percent Change as percent
            Range("L" & SummaryTableRow).Value = FormatPercent(PercentChange)
        
        'Start next summary table row and Opening Price Row
            SummaryTableRow = SummaryTableRow + 1
            OpenPriceRow = i + 1
        'Reset total stock volume
            TotalStockVolume = 0
            
            Else:
        'If stock ticker name is the same, add total stock volume
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
            End If
            
        Next i
        
    Next
          
End Sub
