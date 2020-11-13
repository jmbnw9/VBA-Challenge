Sub Ticker_2016()
    'Set variable for the ticker
    Dim Ticker_Name As String
    
    'Define start
    Dim Start As String
    
    'Set a variable for the yearly change
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    'Set a variable for the percent changed
    Dim Percent_Change As Double
    Percent_Change = 0
    
    'Set variable for the total volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
      
    'Where to place the ticker names
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim start_val As Long
    start_val = 2
    
    Dim rowcount As Long
    rowcount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through all ticker names
    For i = 2 To rowcount
    
        'Check to see if we are still in the same name
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            If Cells(start_val, 3).Value = 0 Then
                For find_val = start_val To i
                    If Cells(find_val, 3).Value <> 0 Then
                        start_val = find_val
                    End If
                Next find_val
            End If
            
            ' Set the name
            Ticker_Name = Cells(i, 1).Value
            
            'Set the yearly change
            Yearly_Change = Cells(i, 6).Value - Cells(start_val, 3).Value
            
            'Set the percent change
            Percent_Change = Round((Yearly_Change / Cells(start_val, 3) * 100), 2)
            
            'Add the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            'Print the ticker name in this column
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            
            'Print the yearly change in this column
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Print the percent change in this column
            Range("K" & Summary_Table_Row).Value = Percent_Change
            
            'Print the stock volume in this column
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset the Total Stock Volume
            Total_Stock_Volume = 0
            
            start_val = i + 1
        
        Else
    
            'Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
        End If
    
    Next i

        
End Sub



