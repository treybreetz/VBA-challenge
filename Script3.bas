Attribute VB_Name = "Module3"
Sub Stocks()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

Dim Ticker As String
Dim yearly_change As String
Dim percent_change As String
Dim total_volume As Double
Dim last_row As Long
Dim summary_row As Long
Dim start_price As Double
Dim end_price As Double

Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double
Dim max_increase_stock As String
Dim max_decrease_stock As String
Dim MaxVolumeStock As String


    summary_row = 2
    start_price = Cells(2, 3).Value
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    max_increase = 0
    max_decrease = 0
    max_volume = 0
    
For i = 2 To last_row

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        end_price = Cells(i, 6).Value
        yearly_change = end_price - start_price
        
    If start_price <> 0 Then
        percent_change = (yearly_change / start_price)
    Else
        percent_change = 0
    End If
    
        total_volume = WorksheetFunction.SumIf(Range("A2:A" & last_row), Ticker, Range("G2:G" & last_row))
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        Cells(summary_row, 9).Value = Ticker
        Cells(summary_row, 10) = yearly_change
        Cells(summary_row, 11) = percent_change
        Cells(summary_row, 12) = total_volume
        
        If percent_change > max_increase Then
            max_increase = percent_change
            max_increase_stock = Ticker
        ElseIf percent_change < max_decrease Then
            max_decrease = percent_change
            max_decrease_stock = Ticker
        End If
        
        If total_volume > max_volume Then
            max_volume = total_volume
            MaxVolumeStock = Ticker
        End If
        
    
        
        summary_row = summary_row + 1
        
        start_price = Cells(i + 1, 3).Value
    End If
Next i
        Range("K2:K" & summary_row).NumberFormat = "0.00%"
        Range("J2:J" & summary_row).NumberFormat = Number

  For j = 2 To summary_row - 1
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
        
           If Cells(j, 11).Value >= 0 Then
            Cells(j, 11).Interior.ColorIndex = 4
        Else
            Cells(j, 11).Interior.ColorIndex = 3
        End If
    Next j
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(2, 16).Value = max_increase_stock
    Cells(3, 16).Value = max_decrease_stock
    Cells(4, 16).Value = MaxVolumeStock
    Cells(1, 17).Value = "Value"
    Cells(2, 17).Value = max_increase
    Cells(3, 17).Value = max_decrease
    Cells(4, 17).Value = max_volume
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
  
    
Next ws
    
    
End Sub
