Sub stock_summary()

    ' Identify variables
    
    Dim Tick_Name As String
    
    Dim Year_Change As Double
    Year_Change = 0
    
    Dim Percent_Change As Double
    Percent_Change = 0
    
    Dim Open_Price As Double
    Open_Price = 0
    
    Dim Close_Price As Double
    Close_Price = 0
    
    Dim Total_Vol As Double
    Total_Vol = 0
    
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    'Determine the last row
    Dim Last_Row As Long
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Format Percent Change Column
    Range("K2:K" & Last_Row).NumberFormat = "0.00%"
    
    Dim start As Long
    start = 2

    'Label Column Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    

 
    'Loop through all ticket symbols
    
    For i = 2 To Last_Row
        
        'Check to see if we are still in the same ticker name
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Add Ticket Symbol to Tick_Name
                Tick_Name = Cells(i, 1).Value
                
                'Add to Total_Vol
                Total_Vol = Total_Vol + Cells(i, 7).Value
                
                'Set Open_Price and Close Price
                Open_Price = Cells(start, 3)
                
                'MsgBox "Open Price is" & Open_Price
                
                Close_Price = Cells(i, 6)
                
                'MsgBox "Close Price is" & Close_Price
                
                'Calculate Year Change
                Year_Change = Close_Price - Open_Price
                
                'Calculate Percent Change
                If Open_Price = 0 Then
                        Percent_Change = 0
                    Else
                        Percent_Change = Year_Change / Open_Price
                    End If
                    
            
                'Add ticker name to Summary Section
                Range("I" & Summary_Table).Value = Tick_Name
                
                'Add total volume to Summary Section
                Range("L" & Summary_Table).Value = Total_Vol
                
                'Add Year Change to Summary Section
                Range("J" & Summary_Table).Value = Year_Change
                    
                    'Conditional Formatting of Year_Change
                    If Year_Change > 0 Then
                        Range("J" & Summary_Table).Interior.Color = RGB(0, 200, 0)
                        
                    ElseIf Year_Change < 0 Then
                        Range("J" & Summary_Table).Interior.Color = RGB(235, 0, 0)
                    
                    End If
                
                'Add Percent Change to Summary Section
                Range("K" & Summary_Table).Value = Percent_Change
            
            
                'Add row to Summary Section
                Summary_Table = Summary_Table + 1
                
                'Reset Total Volume
                Total_Vol = 0
                
                'Reset Open Price
                Open_Price = 0
                
                'Reset Close Price
                Close_Price = 0
                
                'Reset Year Change
                Year_Change = 0
                
                'Reset Percent Change
                Percent_Change = 0
                
                'Sets start the row number of the 1st row of the next Tick Name
                start = i + 1
            
            'If next stock ticker is same as previous cell, keep adding to total
            Else
            
                Total_Vol = Total_Vol + Cells(i, 7).Value
        
        End If
    
    Next i

End Sub

