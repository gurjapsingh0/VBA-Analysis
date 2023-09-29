VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    Dim SummaryRow As Long
    
    ' Loop through each year
    For Each ws In ThisWorkbook.Sheets
        ' Initialize variables for the summary table
        Row1 = 2 ' Start from row 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        GreatestVolumeTicker = ""
        
        ' last row
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through the data
        For i = 2 To LastRow
            ' Check if the current row has the same Ticker as the previous row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Calculate Yearly Change
                YearlyChange = ClosingPrice - OpeningPrice
                
                ' Avoid division by zero error
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Output results to the summary table
                ws.Cells(Row1, 12).Value = Ticker
                ws.Cells(Row1, 13).Value = YearlyChange
                ws.Cells(Row1, 14).Value = PercentChange
                ws.Cells(Row1, 15).Value = TotalVolume
                
                ' Check if this is the first data point for the year
                If YearlyChange <> 0 Then
                    ' Check if this is the first record for the year
                    If GreatestIncrease = 0 Then
                        GreatestIncrease = PercentChange
                        GreatestDecrease = PercentChange
                        GreatestVolume = TotalVolume
                    End If
                    
                    ' Check for greatest % increase, % decrease, and total volume
                    If PercentChange > GreatestIncrease Then
                        GreatestIncrease = PercentChange
                        GreatestIncreaseTicker = Ticker
                    ElseIf PercentChange < GreatestDecrease Then
                        GreatestDecrease = PercentChange
                        GreatestDecreaseTicker = Ticker
                    End If
                    
                    If TotalVolume > GreatestVolume Then
                        GreatestVolume = TotalVolume
                        GreatestVolumeTicker = Ticker
                    End If
                End If
                
                ' Reset variables for the next stock
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                Row1 = Row1 + 1
                TotalVolume = 0
            End If
            
            ' Accumulate Total Volume for the current stock
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Store the Closing Price for the current stock
            ClosingPrice = ws.Cells(i, 6).Value
            
            ' Store the Opening Price for the current stock if it's the first day of the year
            If OpeningPrice = 0 Then
                OpeningPrice = ws.Cells(i, 3).Value
            End If
        Next i
        
        ' Calculate for the last stock of the year
        YearlyChange = ClosingPrice - OpeningPrice
        If OpeningPrice <> 0 Then
            PercentChange = (YearlyChange / OpeningPrice) * 100
        Else
            PercentChange = 0
        End If
        ws.Cells(Row1, 12).Value = Ticker
        ws.Cells(Row1, 13).Value = YearlyChange
        ws.Cells(Row1, 14).Value = PercentChange
        ws.Cells(Row1, 15).Value = TotalVolume
        
        ' Output greatest values
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(2, 18).Value = "Greatest % Decrease"
        ws.Cells(2, 19).Value = "Greatest Total Volume"
        
        ws.Cells(3, 17).Value = GreatestIncreaseTicker
        ws.Cells(3, 18).Value = GreatestDecreaseTicker
        ws.Cells(3, 19).Value = GreatestVolumeTicker
        ws.Cells(3, 20).Value = "Ticker"
        
        ' Labels with location
        ws.Cells(3, 18).NumberFormat = "0.00%"
        
        ws.Cells(4, 20).Value = "Value"
        ws.Cells(4, 17).Value = GreatestIncrease
        ws.Cells(4, 18).Value = GreatestDecrease
        ws.Cells(4, 19).Value = GreatestVolume
        
        ' Clear variables
        Ticker = ""
        OpeningPrice = 0
        ClosingPrice = 0
        TotalVolume = 0
    Next ws
End Sub


