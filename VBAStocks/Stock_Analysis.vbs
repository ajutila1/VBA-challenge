Attribute VB_Name = "Module1"
Sub Stock_Analysis():

'Loop through worksheets
For Each ws In Worksheets

    'Set summary table headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Set variable type for finding last row in dynamic columns
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Set variable type for ticker symbol
    Dim Ticker_Symbol As String

    'Set variable type for total volume and set to 0
    Dim Volume_Total As LongLong
        Volume_Total = 0

    'Set Summary Table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Set first open amount and last close amount variable type
    Dim Open_Number As Double
    Dim Close_Number As Double

    'Grab initial opening value
    Open_Number = Cells(2, 3).Value

    'Loop through all ticker symbols
    For ColumnA = 2 To LastRow

        'Add to volume total
        Volume_Total = Volume_Total + Cells(ColumnA, 7).Value
    
        'Section off searches by ticker symbol
        If Cells(ColumnA + 1, 1).Value <> Cells(ColumnA, 1).Value Then
            
            'Set ticker symbol
            Ticker_Symbol = Cells(ColumnA, 1).Value
        
            'Grab Closing Value
            Close_Number = Cells(ColumnA, 6).Value
        
            'Print ticker symbols in summary table
            Range("I" & Summary_Table_Row).Value = Ticker_Symbol
        
            'Print year change in summary table
            Range("J" & Summary_Table_Row).Value = Close_Number - Open_Number
    
            'Print percent change in summary table, using if conditional to account for 0s in one problem section of data
            If Open_Number = 0 Then
                Range("K" & Summary_Table_Row).Value = (Close_Number - Open_Number)
            Else
                Range("K" & Summary_Table_Row).Value = (Close_Number - Open_Number) / Open_Number
            End If
        
            'Print Volume Total
            Range("L" & Summary_Table_Row).Value = Volume_Total
    
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
            'Reset Volume Total
            Volume_Total = 0
        
            'Grab Opening Number for next loop
            Open_Number = Cells(ColumnA + 1, 3).Value
    
        End If
    
    Next ColumnA

    'Change percent change values in summary table to percent format
    Range("K:K").NumberFormat = "0.00%"

    'Set conditions for color coding yearly change values
    For j = 2 To Summary_Table_Row - 1
            
        If Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
           
    Next j

Next ws

End Sub


