Attribute VB_Name = "Module1"
Sub LoopThruWkb()
'Set your worksheet variable
Dim ws As Worksheet
Application.ScreenUpdating = False
    
'Create a for each loop
For Each ws In Worksheets
ws.Select
        
'Run your macros
Call hw2_moderate
Next
Application.ScreenUpdating = True
End Sub

Sub hw2_moderate()
'Print summary table headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

   
'Set initial variables for holding the ticker name, yearly change, percent change, and total volume
Dim ticker As Variant
Dim yearly_change As Variant
Dim percent_change As Variant
Dim total_vol As Variant
Dim StockOpen As Variant
StockOpen = Cells(2, 3).Value
Dim StockClose As Variant
Dim jCell As Object


total_vol = 0

'Determine last row
Dim last_row As Variant
last_row = Cells(Rows.Count, 2).End(xlUp).Row
        
'Keep track of the location for each ticker
Dim Summary_Table_Row As Variant
Summary_Table_Row = 2
    
'Loop through all tickers
For i = 2 To last_row

    'If the stock open value is not 0 then do the following
    If StockOpen <> 0 Then
            
        'Check if you are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
                            
            'Add the total volume
            total_vol = total_vol + Cells(i, 7).Value
            
            'Set the stock close and yearly change variables
            StockClose = Cells(i, 6).Value
            yearly_change = StockClose - StockOpen
            
            'Add percent change
            percent_change = (yearly_change / StockOpen)
            
            'Print the ticker name, yearly change, percent change, and total volume in the summary table
            Range("I" & Summary_Table_Row).Value = ticker
            Range("J" & Summary_Table_Row).Value = yearly_change
            Range("K" & Summary_Table_Row).Value = percent_change
            Range("L" & Summary_Table_Row).Value = total_vol
            
            'Format cells in column K as a percent
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
                'Format color in column J
                For Each jCell In Range("J" & Summary_Table_Row)
                    If jCell.Value > 0 Then
                        jCell.Interior.Color = vbGreen
                    ElseIf jCell.Value < 0 Then
                        jCell.Interior.Color = vbRed
                    End If
                    Next
                    
            'Add a row to the summary table
            Summary_Table_Row = Summary_Table_Row + 1
                    
            'Reset the total volume and open stock values
            total_vol = 0
            StockOpen = Cells(i + 1, 3).Value
    
            Else
                    
            'Add to the total volume
            total_vol = total_vol + Cells(i, 7).Value
            
        End If
    End If

Next i
    
End Sub




