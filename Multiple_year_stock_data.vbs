Attribute VB_Name = "Module2"
Sub Ticker_Summary_Loop()

    ''### Worksheet Loop code example researched and snippets used below https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

    ' Declare Current as a worksheet object variable.
    
    Dim Current As Worksheet
    
    ' Loop through all of the worksheets in the active workbook.
    
    For Each Current In Worksheets
        
         ' Activate Current Worksheet
         
        Current.Activate
         
         ' Set an initial variable for holding the ticker symbol
         
        Dim ticker_symbol As String
         
         ' Set an initial variable for holding the summary table row
         
        Dim Summary_Table_Row As Integer
        
        ' Set Summary_Table_Row to 2, below headers
        
        Summary_Table_Row = 2
        
        ' Set an initial variable for holding the stock volume total - string due to size of numbers
         
        Dim Stock_Volume As String
        
        Stock_Volume = 0
        
        ' Set an initial variable for holding the yearly change total, yearly change open figure and percent change
        
        Dim Yearly_Change As Double
        
        Yearly_Change = 0
        
        Dim Yearly_Change_Open As Double
        
        Yearly_Change_Open = 0
        
        Dim Percent_Change As Double
        
        ' Set row counter variable - to allow reset of yearly change and percent change between tickers
        
        Dim count1 As Integer
        
        count1 = 1
        
        ' Loop through all the tickers starting after headers
            
            'Find last row of table
            
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            
            'Loop through each row
            
            For i = 2 To lastrow
            
                ' Check for the first stock ticker and set yearly open price, add first ticker row volume to stock volume total
                
                If count1 = "1" Then
                
                    Yearly_Change_Open = Cells(i, 3).Value
                    
                    count1 = count1 + 1
                    
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                
                ' Check if next ticker is the same as the current
                
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                    ' Print the ticker in the Summary Table
                    
                    Range("I" & Summary_Table_Row).Value = ticker_symbol
                    
                    ' Calculate yearly change and percent change, print to summary
                    
                    Yearly_Change = Cells(i, 6).Value - Yearly_Change_Open
    
                    Range("J" & Summary_Table_Row).Value = Yearly_Change
                    
                    Percent_Change = Cells(i, 6).Value / Yearly_Change_Open
                    
                    Range("K" & Summary_Table_Row).Value = Percent_Change - 1
                    
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                    Range("L" & Summary_Table_Row).Value = Stock_Volume
                    
                    ' Move to next summary table row and reset for next ticker
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    Yearly_Change = 0
                    
                    Stock_Volume = 0
                    
                    count1 = 1
                    
                Else
                    
                    ' If next ticker is the same as current, add to count and add volume to stock volume total
                    ticker_symbol = Cells(i, 1).Value
                    
                    count1 = count1 + 1
                    
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                
                End If
                
            Next i
            
            
            
        '%%%%%%%%%%% Start bonus
       
'-------------------------------------------------------------------------------------------------------------------------------------
            'find largest $ change
            Dim first_change As Double
            
            Dim change_ticker As String
            
            first_change = Cells(2, 10).Value
            
            change_ticker = Cells(2, 9).Value
            
            lastrow = Cells(Rows.Count, 10).End(xlUp).Row
            
            'Loop through each row
            For i = 2 To lastrow
            
                 If Cells(i, 10).Value > first_change Then
                 
                 first_change = Cells(i, 10).Value
                 
                 change_ticker = Cells(i, 9).Value
                 
                 End If
                 
            Next i
        
            Cells(1, 17).Value = first_change
            
            Cells(1, 16).Value = change_ticker
            
'-------------------------------------------------------------------------------------------------------------------------------------
            
            ' Find largest % change
            
            Dim first_change2 As Double
            
            Dim change_ticker2 As String
            
            first_change2 = Cells(2, 11).Value
            
            change_ticker2 = Cells(2, 9).Value
            
            lastrow = Cells(Rows.Count, 11).End(xlUp).Row
            
            'Loop through each row
            For i = 2 To lastrow
            
                 If Cells(i, 11).Value > first_change2 Then
                 
                 first_change2 = Cells(i, 11).Value
                 
                 change_ticker2 = Cells(i, 9).Value
                 
                 End If
                 
            Next i
        
            Cells(2, 17).Value = first_change2
            
            Cells(2, 16).Value = change_ticker2
            
'-------------------------------------------------------------------------------------------------------------------------------------

            ' Find Largest Trading Volume
            
            Dim first_change3 As String
            
            Dim change_ticker3 As String
            
            first_change3 = Cells(2, 12).Value
            
            change_ticker3 = Cells(2, 9).Value
            
            lastrow = Cells(Rows.Count, 12).End(xlUp).Row
            
            'Loop through each row
            
            For i = 2 To lastrow
            
                 If Cells(i, 12).Value > first_change3 Then
                 
                 first_change3 = Cells(i, 12).Value
                 
                 change_ticker3 = Cells(i, 9).Value
                 
                 End If
                 
            Next i
        
            Cells(3, 17).Value = first_change3
            
            Cells(3, 16).Value = change_ticker3
              
        ' Alert user to completion of ticker summary for each year
        
        MsgBox ("Sheet " + Current.Name + " Summary Is Ready")
         
    Next
    
End Sub



