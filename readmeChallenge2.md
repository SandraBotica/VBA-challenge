readmeChallenge
Sub ticker():

            'Inserting titles of columns Via cells for the bonus challenge on the first worksheet only
    Cells(1, 16).Value = "Ticker "
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest_Percent_Increase"
    Cells(3, 15).Value = "Greatest_Percent_Decrease"
    Cells(4, 15).Value = "Total_Stock_Volume"

            'Applying script to data set on every worksheet
For Each ws In Worksheets

            'Variable declarations and variable assignments
            'Set and initial variable for holding the ticker symbol name
    Dim ticker_symbol As String
    
            'Set and initial variable for holding the total stock volume number
    Dim total_stock_volume As LongLong
    total_stock_volume = 0
    
            'Keep track of the location for each variable in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
            'Set and initial variable for holding the open price number for a ticker
    Dim open_price As Double
    
            'Set and initial variable for holding the close price number for a ticker
    Dim close_price As Double
    
            'Set and initial variable for holding the Yearly_Change number
    Dim Yearly_Change As Double
    
            'Set and initial variable for holding the Percent_Change number
    Dim Percent_Change As Double
    
            'Set and initial variable for holding the Percent_Change_rounded number
    Dim Percent_Change_rounded As Double
    
            'Set and initial variable for holding the Greatest_Percent_Increase number for the Bonus summary table
    Dim Greatest_Percent_Increase As Double
    
            'Set and initial variable for holding the Greatest_Percent_Decrease number for the Bonus summary table
    Dim Greatest_Percent_Decrease As Double
    
            'Set and initial variable for holding the Greatest_total_volume number for the Bonus summary table
    Dim Greatest_total_volume As LongLong
    
            'Determine the lastrow
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
            'Inserting titles of columns Via cells on every worksheet
    ws.Cells(1, 9).Value = "Ticker "
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
    
            'Loop through all rows
    For i = 2 To lastrow
    
                    'Variable argument
         total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
         
                    'Variable argument for open price for the first ticker to assist with calculating Yearly_Change and Percent_Change
         open_price = ws.Cells(i, 3).Value
         
                    'Check if ticker symbol from one row to the next is the same'
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    'Variable argument to identify ticker symbol
            ticker_symbol = ws.Cells(i, 1).Value
            
                    'Put the ticker symbol in a row of the summary table
            ws.Range("I" & summary_table_row).Value = ticker_symbol
                    
                    'Put the total stock volume for each of the rows with this ticker value into this cell
            ws.Range("L" & summary_table_row).Value = total_stock_volume
            
                    'Variable argument for close price for the first ticker to assist with calculating Yearly_Change and Percent_Change
         close_price = ws.Cells(i, 6).Value
            
                    'Variable argument for Yearly_Change for the first ticker
            Yearly_Change = (close_price - open_price)
            
                'Print the ticker_symbol into the summary table
            ws.Range("I" & summary_table_row).Value = ticker_symbol
            
                'Print the Yearly_Change amount into the summary table
            ws.Range("J" & summary_table_row).Value = Yearly_Change
            
                
                
                
                
             If ws.Yearly_Change = Positive Then
                ws.Yearly_Change.Interior.ColorIndex = 3
            Else
                If ws.Yearly_Change.Value = Negative Then
                ws.Yearly_Change.Interior.ColorIndex = 5
            End If
                
                
                
                    'Variable argument for Percent_Change for the first ticker
            Percent_Change = (Yearly_Change / open_price) * 100
            
                    'Print the ticker_symbol into the summary table
            ws.Range("I" & summary_table_row).Value = ticker_symbol
                        
                    'Print the Percent_Change amount into the summary table
            ws.Range("K" & summary_table_row).Value = Percent_Change
            
            
            open_price = ws.Cells(i + 1, 3).Value
            
            
            
            Percent_Change_rounded = Round(Percent_Change)
            
            
                        'Add one to the summary table row to move to the next row and new ticker
            summary_table_row = summary_table_row + 1
            
                        'Reset the total_stock_volume, Yearly_Change and Percent_Change total to zero
            total_stock_volume = 0
            Yearly_Change = 0
            Percent_Change = 0
            
            End If
            
        Next i
        
            
Next ws

End Sub
