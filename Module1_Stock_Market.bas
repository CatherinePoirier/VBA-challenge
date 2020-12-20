Attribute VB_Name = "Module1"
Sub StockMarket():

Dim ws As Worksheet ' defines integer to run on all worksheet

 For Each ws In ThisWorkbook.Worksheets 'Begin doing the below on all workbooks sheets



Dim Ticker As String  'Column B or 2
Dim Open_Stock_price As Double ' for price of first of year
Dim Close_Stock_price As Double ' for rprice of Close Stock
Dim Change_in_Price As Double
Dim i As Long
Dim last_row_1 As Double
Dim Percent_change As Double
Dim Total_Volume As Double
Dim Stock_Volume As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim last_row_9 As Double
Dim Ticker_Decrease As String ' Ticker for Greatest Decrease
Dim Ticker_Increase As String
Dim Ticker_Volume As String
Dim i_bonus As Integer


Dim i_summary As Integer

    'Insert Headings
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"  'change in price header
    ws.Cells(1, 11) = "Percent Change Yearly"
    ws.Cells(1, 12) = "Total Stock Volume Yearly"
    'ws.Cells(1, 18) = "Opening Price"  'remove this later not needed for homework
    'ws.Cells(1, 19) = "Closing Price"  'remove this later not needed for homework


    last_row_1 = Cells(Rows.Count, 1).End(xlUp).Row  'Finds the last row in a column
    i_summary = 2  ' Counts Rows for summary table
    Open_Stock_price = ws.Cells(2, 3)  'Assigns Initial opening Price
    Stock_Volume = 0
    

'iterate through all rows
   
   
   For i = 2 To last_row_1
    
    Ticker = ws.Cells(i, 1).Value  'Assigns Ticker to the first ticker symbol
    Close_Stock_price = ws.Cells(i, 6) ' Grabs Close Stock Price t
             
             If Ticker <> ws.Cells(i + 1, 1).Value Then   'conditional when Row value for Column One changes then we have a new ticker
        
             Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value ' adds each row of Stock Volume
             
             'Assigns Values to Summary Table
              ws.Cells(i_summary, 12).Value = Stock_Volume ' Assigns value to summary table
              ws.Cells(i_summary, 9).Value = Ticker  'Assigns Ticker value to summary table
              'ws.Cells(i_summary, 19).Value = Close_Stock_price  ' REMOVE later for testing
              'ws.Cells(i_summary, 18) = Open_Stock_price  ' REMOVE later for testing
             
             Change_in_Price = Close_Stock_price - Open_Stock_price  'Calculates Change_In_Price
                        
        
             ws.Cells(i_summary, 10).Value = Change_in_Price 'Assigns Change_in_Price value to summary table Yearly Change
              
                    If Open_Stock_price = 0 Then
                        Percent_change = 0
                        Else
                        'Percent_change = Round((Change_in_Price / Open_Stock_price), 2) * 100
                        Percent_change = ((Change_in_Price / Open_Stock_price) * 100)
                    End If
                                                      
                ' Formats Percent Change and adds to summary table
                ws.Cells(i_summary, 11) = "%" & Percent_change
              
                
                

                        
                        'Format color to Red if negative for Change_in_price
                       If ws.Cells(i_summary, 10) < 0 Then
                            ws.Cells(i_summary, 10).Interior.ColorIndex = 3 ' Changes Cells to Red
                            ws.Cells(i_summary, 10).Font.ColorIndex = 1
                            ws.Cells(i_summary, 10).Font.ColorIndex = 1
                            ws.Cells(i_summary, 10).NumberFormat = "0.00"
                          'Formats and Green if positive for Change_in_price
                        Else
                              ws.Cells(i_summary, 10).Interior.ColorIndex = 4 ' Changes cells to Green
                              ws.Cells(i_summary, 10).Font.ColorIndex = 1
                              ws.Cells(i_summary, 10).NumberFormat = "0.00"
                        End If
                
              i_summary = i_summary + 1
              
        'Defines Open Stock Price for next Tickerand resets Stock Volume
          Open_Stock_price = ws.Cells(i + 1, 3)
          Stock_Volume = 0
          
     'Continues to add up a Tickers Stock_Volume
       Else
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value 'Adds last row of Stock Volume Price
                     
              '
                        
       End If
           Next i

' search summary table for greatest Percent Increase column 11
    'greatest % decrease column 11
    'greatest Total Volume column 12
    'ticker is column 9
    'define headers
    'define variables greatest_increase, greatest_decrease greatest_volume

last_row_9 = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox (last_row_9)

    greatest_decrease = 0
    i_bonus = 2
 ' MsgBox (greatest_decrease)
  
'Insert Headings
    ws.Cells(1, 15) = "Ticker"
    
    ws.Cells(1, 16) = "Value"
    ws.Cells(3, 14) = "Greatest % Decrease"
    ws.Cells(2, 14) = "Greatest % Increase"
    ws.Cells(4, 14) = "Greatest Total Volume"
    ws.Columns("N").ColumnWidth = 20
    ws.Columns("P").ColumnWidth = 15

greatest_decrease = ws.Cells(2, 11)
greatest_increase = ws.Cells(2, 11)
greatest_volume = ws.Cells(2, 12)
Ticker_Decrease = ws.Cells(2, 9)
Ticker_Increase = ws.Cells(2, 9)
Ticker_Volume = ws.Cells(2, 9)

    For n = 2 To last_row_9
    
        'Finds Greatest_Decrease and assigns value and Ticker
       If greatest_decrease > ws.Cells(n, 11).Value Then
            greatest_decrease = ws.Cells(n, 11).Value
            Ticker_Decrease = ws.Cells(n, 9).Value  'Assigns Ticker
    
        End If
            
        'Finds Greatest_Increase and assigns value and Ticker
        If greatest_increase < ws.Cells(n, 11).Value Then
            greatest_increase = ws.Cells(n, 11).Value
            Ticker_Increase = ws.Cells(n, 9).Value  'Assigns Ticker
       End If
            
      'Finds Greatest_Volume and assigns value and Ticker
        If greatest_volume < ws.Cells(n, 12).Value Then
            greatest_volume = ws.Cells(n, 12).Value
            Ticker_Volume = ws.Cells(n, 9).Value  'Assigns Ticker
       End If
          
            
        Next n
        greatest_increase = (greatest_increase * 100)
         greatest_decrease = (greatest_decrease * 100)
        
     'Assigns Values to Summary Table
    ws.Cells(i_bonus, 16) = "%" & greatest_increase
         'ws.Cells(i_bonus, 16).NumberFormat = "0.00"  ' formats to 2 decimals
    ws.Cells(i_bonus, 15) = Ticker_Increase
    ws.Cells(i_bonus + 1, 16) = "%" & greatest_decrease
    'ws.Cells(i_bonus + 1, 16).NumberFormat = "0.00"  ' formats to 2 decimals
    ws.Cells(i_bonus + 1, 15) = Ticker_Decrease
    
    ws.Cells(i_bonus + 2, 16) = greatest_volume
    ws.Cells(i_bonus + 2, 15) = Ticker_Volume
    Columns("N").ColumnWidth = 20
'MsgBox ("Done with first worksheet going to second")
    
 Next ws
'MsgBox ("second worksheet")
End Sub



