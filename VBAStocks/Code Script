Sub Stock()
    
    On Error Resume Next
    
    'declare ticker as string
    Dim Ticker As String
    
    'declare lastrow to navigate
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'declare variables as double
    Dim Yearly_change As Double
        Yearly_change = Ychange
        
  
    Dim Percent_change As Double
        Percent_change = Pchange
        
    'declare variable for format
        Pchange = Selection
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
        
              
    Dim Total_Stock_Volumn As Double
        Total_Stock_Volumn = Volumn
    
    'delcare the table as integer
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
 
    
    'set the initial variable value
    Open_price = Cells(2, 3).Value
  
      'set i as row count
      For i = 2 To lastRow
     
         'if statement when the next row cell is different than last row in first column
         'then the ticker will show the value of next different cell's value in the list
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
            'total volumn is equal to the same ticker added up value
            Volumn = Volumn + Cells(i, 7).Value
            
            'set close price in the column 6
            Close_price = Cells(i, 6).Value
            
            
            'yearly change is equal to close price at the end of the year - open price at the beginning of the year
            Ychange = Close_price - Open_price
             
                'percent change is yealy change divided by open price
                If Open_price <> 0 Then
                
                    Pchange = Ychange / Open_price
            
                Else
                    ipchange = 100
             End If
                
            Open_price = Cells(i + 1, 3).Value
            
            'set each columns to run corresponding variables
            Range("K" & Summary_Table_Row).Value = Ticker
            
            Range("L" & Summary_Table_Row).Value = Ychange
            
            Range("M" & Summary_Table_Row).Value = Pchange
            
            Range("N" & Summary_Table_Row).Value = Volumn
             
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Total volumn calculation is incremental which needs to set 0 to add up to previous row until next ticker
            Volumn = 0

            
            Else
            
                 
                'otherwise the volumn is still calculated as such
                Volumn = Volumn + Cells(i, 7).Value

     End If
    
   Next i
    
    'set rowcount for column 12 specifically to change format based on values of each cells
    RowCount = Cells(Rows.Count, 12).End(xlUp).Row
    
        'set j as new variable to specifically declare change for column in yearly change
        For j = 2 To RowCount
        
        'set yearly change column
        Ychange = Cells(j, 12).Value
            
            'if statement when the value is positive or 0 then its Green otherwise its Red
            If Ychange >= 0 Then
         
                Cells(j, 12).Interior.ColorIndex = 4
        
            Else
            
                Cells(j, 12).Interior.ColorIndex = 3
                
            End If
   
   Next j
        
        'declare variables which will represent the function of the value
        Dim Max As Double
        Dim Min As Double
        Dim Vol As Double
        
       
        'declare % format cells
        With Range("Q2", "Q3").Select
        
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.0%"
        Selection.NumberFormat = "0.00%"
        
        
        End With
         
           'Maximum % change in column M function to its last row in the range
           Max = Application.WorksheetFunction.Max(Range("M2:M" & lastRow))
            
           'the below returns row number, from which the ticker column is
           i = Application.WorksheetFunction.Match(Max, Range("M2:M" & lastRow), 0)
        
           'delcare the cell to show the value
            Cells(2, 17).Value = Max
            
           'declare the cell to show which row the ticker value is residing
            Cells(2, 16).Value = Cells(i, 11).Value
            
            
            
            'Minimum % change in column M function to its last row in the range
            Min = Application.WorksheetFunction.Min(Range("M2:M" & lastRow))
            
            'the below returns row number, from which the ticker column is
            i = Application.WorksheetFunction.Match(Min, Range("M2:M" & lastRow), 0)
            
            'declare the cell to show the value
            Cells(3, 17).Value = Min
            
            'declare the cell to show the ticket value which row is residing
            Cells(3, 16).Value = Cells(i, 11).Value
            
            
            
            'Maximum total volumn in column N function to its last row in the range
            Vol = Application.WorksheetFunction.Max(Range("N2:N" & lastRow))
            
            'the below returns row number, from which the ticker column is
            i = Application.WorksheetFunction.Match(Vol, Range("N2:N" & lastRow), 0)
          
            'delcare the cell to show the value
            Cells(4, 17).Value = Vol
         
            'delcare the cell to show the ticker value which row is risiding
            Cells(4, 16).Value = Cells(i, 11).Value
           


End Sub

