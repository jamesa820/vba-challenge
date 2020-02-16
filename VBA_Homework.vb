
Sub Stock_Challenge()

        'Determine and assign variables
Dim Ticker_value As String

Dim Table As Integer

        'opening value of each stock for the year
Dim Begin_year_open As Double

        'closing value of each stock for that year
Dim End_yr_close As Long

        'Variable to count the rows to scan data all the way to the bottom
Dim Row_Counts As Long

        'The difference between beginning and end of year
Dim Year_change As Double

        'Total volume of stocks traded per stock for the year
Dim Total_volume As Double

        'The number of stocks traded on the first day of the year
Dim Start_stock_volume As Long

            'Percentage of change
Dim Percentage As Double

            'The number of stocks traded on the last day of the year
Dim End_stock_volume As Long


  'assigning variables for column headers
    Dim ticker As String
    Dim change As String
    Dim percent As String
    Dim total As String
    
 'Assigning initial values to variables
    Table = 2
    Begin_year_open = 0
    End_yr_close = 0
    Year_change = 0
    Percentage = 0
    Start_stock_volume = 0
    End_stock_volume = 0
    Total_volume = 0
    
  'assigning value of variable for new column headers
    ticker = "Ticker"
    change = "Yearly Change"
    percent = "Percent Change"
    total = "Total Stock Volume"
    
  'assigning new column headers to their position
    Cells(1, 9).Value = ticker
    Cells(1, 10).Value = change
    Cells(1, 11).Value = percent
    Cells(1, 12).Value = total
    
 'Change width on column to accomodate headers
 'Change formatting on column to present percentages
    Columns("J").ColumnWidth = 18
    Columns("K").ColumnWidth = 18
    Columns("L").ColumnWidth = 18
    Columns("K").NumberFormat = "0.00%"
    
    
 'creating conditions to allow full rows to be examined
Row_Counts = Cells(Rows.Count, 1).End(xlUp).Row


 'For loop to scan rows
For i = 2 To Row_Counts
    
  '...
    'If Start_stock_volume = 0 Then
        
        'Start_stock_volume = Cells(i + 1, 7).Value
    
    'End If
    
  'Determining opening value of individual stocks
    If Begin_year_open = 0 Then
       
  'assigning variable to row to find year opening value
       Begin_year_open = Cells(i, 3).Value
         

    End If
    
    
  'Finding end of ticker for each stock
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
  'Printing ticker on list
        Ticker_value = Cells(i, 1).Value
        Cells(Table, 9).Value = Ticker_value
        
  'assigning end of year closing value to variable
        End_yr_close = Cells(i, 6).Value
        
  'Finding and printing the difference for open and close for year
        Year_change = End_yr_close - Begin_year_open
        
        
  'Printing the differece in the year change
        Year_change = Cells(Table, 10).Value
        
  'Formula for finding the percentage of change from open to close
        Percentage = (Year_change / Begin_year_open)
        
        Cells(Table, 11).Value = Percentage
  
  
        Total_volume = Total_volume + Cells(i, 7).Value
        
        Cells(Table, 12).Value = Total_volume
       
  'Adding a row to make make sure the table populates in each cell.
        Table = Table + 1
        
  'Reset variables to get accurate counts for each stock
        Begin_year_open = 0
                 
        End_yr_close = 0
        
        Year_change = 0
        
        Percentage = 0
        
        Total_volume = 0
        
        Else
        
        Total_volume = Total_volume + Cells(i, 7).Value
    End If
            
  'New loop to set color indicating conditions for Yearly Change
Next i

  'Created variable for row counts
Dim Row_Counts2 As Long

  'Assigned value to variable
Row_Counts2 = Cells(Rows.Count, 10).End(xlUp).Row

  'Set up For loop
    For J = 2 To Row_Counts2
  'Set conditions for color designation within the row
        If Cells(J, 10) < 0 Then
        
  'assigning colors to cells according to set conditions
        Cells(J, 10).Interior.ColorIndex = 3
            
            Else
    
            Cells(J, 10).Interior.ColorIndex = 4
    
        End If
    Next J

End Sub

   

