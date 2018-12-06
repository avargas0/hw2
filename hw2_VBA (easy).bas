Attribute VB_Name = "Module1"
Sub LoopThruWkb()
    'Set your worksheet variable
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    'Create a for each loop
    For Each ws In Worksheets
        ws.Select
        
        'Run your macros
        Call hw2_easy
        
        'Auto-resize columns
        ws.Columns("I:J").AutoFit
        
    Next
    
    Application.ScreenUpdating = True

End Sub
Sub hw2_easy()
    'Print summary table headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 9).Font.Bold = True
    
    Cells(1, 10).Value = "Total Volume"
    Cells(1, 10).Font.Bold = True
    
    'Set an initial variable for holding the ticker name
    Dim ticker As Variant
    
    'Set an initial variable for holding the total volume
    Dim total_vol As Variant
    total_vol = 0
    
    'Determine last row
    Dim last_row As Variant
        last_row = Cells(Rows.Count, 2).End(xlUp).Row
        
    'Keep track of the location for each ticker
    Dim Summary_Table_Row As Variant
    Summary_Table_Row = 2
    
    'Loop through all tickers
    For i = 2 To last_row
    
        'Check if you are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
        'Set the total volume name
         ticker = Cells(i, 1).Value
        
        'Add the total volume
        total_vol = total_vol + Cells(i, 7).Value
        
        'Print the ticker name in the summary table
        Range("I" & Summary_Table_Row).Value = ticker
        
        'Print the total volume in the summary table
        Range("J" & Summary_Table_Row).Value = total_vol
        
        'Add a row to the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the total volume
        total_vol = 0
    
    'If the cell immediately following a row is the same ticker...
    Else
        
        'Add to the total volume
        total_vol = total_vol + Cells(i, 7).Value
        
    End If
    
    Next i
    
End Sub

