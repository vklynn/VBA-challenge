Sub Stock_Data()
'Loop Through All Sheets
For Each ws In Worksheets

'Set header to Ticker
ws.Cells(1, 9).Value = "Ticker"
'Set header to Yearly Change
ws.Cells(1, 10).Value = "Yearly Change"
'Set header percentage change
ws.Cells(1, 11).Value = "Percent Change"
'Set header to Total Stock Volume
ws.Cells(1, 12).Value = "Total Stock Volume"

    'Create variable to hold the file name
    Dim WorksheetsName As String
    
    'Set initial variable for ticker name
    Dim Ticker_Name As String

    'Set variable foryear change
    'Brand_total in example
    Dim Yearly_Change As Double
    Yearly_Change = 0

    'Set variable for total stock volume
    Dim Stock_Volume As Double
    Stock_Volume = 0

    '% change variables
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim Open_Val As Double
    Open_Val = 0
    Dim Close_Val As Double
    Close_Val = 0
    'Keep track of location of ticker name in column
    Dim Ticker_Table As Integer
    Ticker_Table = 2

    ' Counts the number of rows
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    'Loop to go through all the clicker names
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Set Ticker Name
            Ticker_Name = ws.Cells(i, 1).Value
            
            'Open close test
            Open_Val = Open_Val + ws.Cells(i, 3).Value
            Close_Val = Close_Val + ws.Cells(i, 6).Value
            
            'Set yearly Change
            Yearly_Change = Close_Val - Open_Val
            'Get Percent_Change  ws.Cells(i, 6).Value - ws.Cells(i, 3).Value / ws.Cells(i, 3).Value
            Percent_Change = (Close_Val - Open_Val) / Open_Val
            'Set total stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
            'Print Ticker Name in col
            ws.Range("I" & Ticker_Table).Value = Ticker_Name
            'Print Yearly Change
            ws.Range("J" & Ticker_Table).Value = Yearly_Change
            'Print % change
            ws.Range("K" & Ticker_Table).Value = Percent_Change
            'Print total stock volume
            ws.Range("L" & Ticker_Table).Value = Stock_Volume
        
            'Add ticker to collum
            Ticker_Table = Ticker_Table + 1
            'Reset yearly change to
            Open_Val = 0
            Close_Val = 0
            Yearly_Change = 0
            'Reset Percent_Change
            Percent_Change = 0
            'Reset total stock volume
            Stock_Volume = 0
    
        Else
            'cal open & close val
            Open_Val = Open_Val + ws.Cells(i, 3).Value
            Close_Val = Close_Val + ws.Cells(i, 6).Value
            ' caculate yearly change
            Yearly_Change = Close_Val - Open_Val
            'caculate stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            'Get Percent_Change
            Percent_Change = (Close_Val - Open_Val) / Open_Val
        End If
    Next i
Next ws

MsgBox ("Data Updated")

End Sub
