Attribute VB_Name = "Module1"
Sub StockAnalysisA()
   
    'This code was developed by Carlos Valverde Martínez as part of the week 2 VBA Homework - The VBA of Wall Street.
   
    'Code for Ticker Column. This part of the code is for the first column.
   
    'Header names
     
     Cells(1, 9).Value = "Ticker"
    
    'Defining variables
     
     Dim Ticker As String
    
    'Defining variables for the summary table
     
     Dim Summary_Row As Integer
     
     Summary_Row = 2
        
    'Defining variable and counting the rows for column A.
    
     Dim NumberRows As Double
     NumberRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Checkpoint Number of rows
     'Cells(1, 15).Value = NumberRows
    
    'Loop for the Ticker names summary in column 9 "Ticker"
    
     For I = 2 To NumberRows
    
        If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
              
            Ticker = Cells(I, 1).Value
        
            Cells(Summary_Row, 9).Value = Ticker
                                 
           'Checkpoint Number of rows per ticker
            'Cells(Summary_Row, 15).Value = i
                                 
            Summary_Row = Summary_Row + 1
         
        End If
        
     Next I
    
'___________________________________________________
       
    'Code for Yearly Change Column. This part of the code is for the second column.
        
    'Header names
     
     Cells(1, 10).Value = "Yearly Change"

    'Defining variables
     
     Dim Yearly_Change As Double
     Dim Year_Initial_Value As Double
     Dim Year_Final_Value As Double
    
     Summary_Row = 2
    
    'Defining variable and counting the rows for column C.
    
     Dim NumberRowsValue As Double
     NumberRowsValue = Cells(Rows.Count, 3).End(xlUp).Row
    
    'Loop for the Yaerly change summary in column 10 "Yearly Change"
    
     For I = 2 To NumberRowsValue
    
        If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
        
            Year_Initial_Value = Cells(I, 3).Value
            
            'Checkpoint Initial value
            'Cells(Summary_Row, 12).Value = Year_Initial_Value
                         
        ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            Year_Final_Value = Cells(I, 6).Value
            
            'Checkpoint Final value
            'Cells(Summary_Row, 13).Value = Year_Final_Value
            
            Yearly_Change = Year_Final_Value - Year_Initial_Value
            
            Cells(Summary_Row, 10).Value = Yearly_Change
            
            Summary_Row = Summary_Row + 1
            
        End If
        
      Next I
        
'___________________________________________________
    
    'Code for Percent Change Column. This part of the code is for the third column.
        
     'Header names
    
      Cells(1, 11).Value = "Percent Change"

     'Defining variables
     
      Dim Percent_Change As Double
      Dim Year_Initial_Value_Percent As Double
      Dim Year_Final_Value_Percent As Double
    
      Summary_Row = 2
    
     'Defining variable and counting the rows for column C.
    
      Dim NumberRowsPercent As Double
      NumberRowsPercent = Cells(Rows.Count, 3).End(xlUp).Row
    
     'Loop for the Percent Change summary in column 11 "Percent Change"
    
      For I = 2 To NumberRowsPercent
    
        If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
        
            Year_Initial_Value_Percent = Cells(I, 3).Value
            
            'Checkpoint Initial value
            'Cells(Summary_Row, 12).Value = Year_Initial_Value_Percent
                                              
        ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            Year_Final_Value_Percent = Cells(I, 6).Value
                   
            'Checkpoint Final value
            'Cells(Summary_Row, 13).Value = Year_Final_Value_Percent
                   
            'Next line when there is an error when Year_Initial_Value_Percent = 0
            
            If Year_Initial_Value_Percent = 0 Then
            
                Percent_Change = 0
                
            Else
            
                Percent_Change = Year_Final_Value_Percent / Year_Initial_Value_Percent - 1
                        
            End If
                        
            Cells(Summary_Row, 11).Value = Percent_Change
            
            Cells(Summary_Row, 11).NumberFormat = "0.00%"
            
            Summary_Row = Summary_Row + 1
              
        End If
                                         
      Next I

'___________________________________________________

    'Code for Total Stock Volume Column. This part of the code is for the fourth column.
        
    'Header names
     
     Cells(1, 12).Value = "Total Stock Volume"

    'Defining variables
    
     Dim Total_Stock_Volume As Double
     Total_Stock_Volume = 0
     
     Dim Initial_Stock_Volume As Double
     Initial_Stock_Volume = 0

    'Defining variable and counting the rows for column H.
    
     Dim NumberRowsVolume As Double
     NumberRowsVolume = Cells(Rows.Count, 1).End(xlUp).Row
    
     Summary_Row = 2
    
    'Loop for the Total Stock Volume summary in column 12 "Total Stock Volume"
    'I use variable type as Double or Long instead of Integer y works
                
      For I = 2 To NumberRowsVolume

        If Cells(I, 1).Value = Cells(I + 1, 1).Value Then
               
            Total_Stock_Volume = Initial_Stock_Volume + Cells(I, 7).Value
                               
            Cells(Summary_Row, 12).Value = Total_Stock_Volume
            
            Initial_Stock_Volume = Cells(Summary_Row, 12).Value
                                                              
        ElseIf Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
                                                                 
            Cells(Summary_Row, 12).Value = Total_Stock_Volume + Cells(I, 7).Value
                                                                 
            Summary_Row = Summary_Row + 1
            
            Initial_Stock_Volume = 0
            
            Total_Stock_Volume = 0
            
        End If
                                      
     Next I
    
'___________________________________________________

    'Code to color green the positive change and red the negative change
    
     Dim NumberRowsYearlyChange As Double
     NumberRowsYearlyChange = Cells(Rows.Count, 10).End(xlUp).Row
     
     For I = 2 To NumberRowsYearlyChange
     
        If Cells(I, 10).Value > 0 Then
     
            Cells(I, 10).Interior.ColorIndex = 4
            
        ElseIf Cells(I, 10).Value < 0 Then
        
            Cells(I, 10).Interior.ColorIndex = 3
            
        End If
        
     Next I
            
'___________________________________________________
               
    'Code for bonus Greatest % Increase
    
    'Header names
     
     Cells(1, 16).Value = "Ticker"
     
     Cells(1, 17).Value = "Value"
     
     Cells(2, 15).Value = "Greatest % Increase"
     
     Dim GreatestIncrease As Double
     
     Dim NumberRowsIncrease As Double
     NumberRowsIncrease = Cells(Rows.Count, 11).End(xlUp).Row
         
     GreatestIncrease = Application.WorksheetFunction.Max(Range("K2:K" & NumberRowsIncrease))
     
     Cells(2, 17).Value = GreatestIncrease
     
     Cells(2, 17).NumberFormat = "0.00%"
     
     'With the next line the ticker name is located for Greatest % Increase, with the match command the row is defined.
     
     Cells(2, 16).Value = Cells(Application.WorksheetFunction.Match(GreatestIncrease, Range("K:K"), 0), 9).Value
    
'___________________________________________________
               
    'Code for bonus Greatest % Decrease
    
     Cells(3, 15).Value = "Greatest % Decrease"
     
     Dim GreatestDecrease As Double
     
     Dim NumberRowsDecrease As Double
     NumberRowsDecrease = Cells(Rows.Count, 11).End(xlUp).Row
     
     GreatestDecrease = Application.WorksheetFunction.Min(Range("K2:K" & NumberRowsDecrease))
     
     Cells(3, 17).Value = GreatestDecrease
     
     Cells(3, 17).NumberFormat = "0.00%"
     
     'With the next line the ticker name is located for Greatest % Decrease, with the match command the row is defined.
     
     Cells(3, 16).Value = Cells(Application.WorksheetFunction.Match(GreatestDecrease, Range("K:K"), 0), 9).Value
     
'___________________________________________________
               
    'Code for bonus Greatest % Total Volume
    
     Cells(4, 15).Value = "Greatest Total Volume"
     
     Dim GreatestTotalVolume As Double
     
     Dim NumberRowsTotalVolume As Double
     NumberRowsTotalVolume = Cells(Rows.Count, 12).End(xlUp).Row
               
     GreatestTotalVolume = Application.WorksheetFunction.Max(Range("L2:L" & NumberRowsTotalVolume))
     
     Cells(4, 17).Value = GreatestTotalVolume
     
     'With the next line the ticker name is located for Greatest Total Volume, with the match command the row is defined.
     
     Cells(4, 16).Value = Cells(Application.WorksheetFunction.Match(GreatestTotalVolume, Range("L:L"), 0), 9).Value
                      
End Sub


