

Sub stocks()

'Delcare variables'
    Dim ws As Worksheet
    Dim Worksheet As String
    Dim Ticker_Name As String
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    Dim Yearly_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Variant
    Dim Value As Long
    Dim Close_Price As Double
    Dim J As Long
    Dim Summary_Table_Row As Integer
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Variant

'Set column names for summary table calculations'
For Each ws In Worksheets
    Summary_Table_Row = 2
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    'Find the last row of A column for Tickers'
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Find the last row of K column for Percent Change'
    lastrowK = ws.Cells(Rows.Count, "K").End(xlUp).Row

    
    For i = 2 To lastrow
    'Locate ticker names and paste onto summary table'
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
       
    'Calculate Yearly Change with closing price minus opening price. Paste onto summary table'
        Close_Price = ws.Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
     
        'Conditional formatting for Yearly Change column'
            If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                'Red for negative'
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                'Green for positive'
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            End If
            
         'Conditional formatting for Percent Change'
            If Open_Price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = (Close_Price - Open_Price) / Open_Price
                
            End If

    'Print Percent Change'
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    'Update format of K column to show percentages'
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     
     'Calculate total stock volume using for each symbol'
        Total_Stock_Volume = Cells(i, 7).Value + Total_Stock_Volume
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        Ticker_Name = 0
        Total_Stock_Volume = 0
        Greatest_Increase = 0
    Else

        Ticker = ws.Cells(i, 1).Value
        OpeningPrice = ws.Cells(i, 3).Value
    End If

    Next i
    
    
    'Create headers for summary table and let each value equal zero'
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    Greatest_Increase = 0
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    Greatest_Decrease = 0
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    Greatest_Total_Volume = 0
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    
    'Find last row for the Summary Table '
    lastrow_summary_table = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For i = 2 To lastrow_summary_table
            
            'Calculate Greatest % Increase and format as a percentage'
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Calculate Greatest % Decrease and format as a percentage'
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"

            'Calculate Greatest Total Volume'
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value

            End If
        
        Next i
Next ws


End Sub

