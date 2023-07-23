Sub Stock_Market_Analysis():

' The below script will run an analysis on the generated stock market data and produce two summary tables with the findings.

  ' The variables referenced to multiple times in this script are declared below:
  Dim ws As Worksheet

  Dim Ticker_Symbol As String
  Dim Yearly_Change As Double
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  Dim Percent_Change As Double
  Dim Total_Stock_Volume As Double

  Dim i As Double
  Dim j As Double
  Dim Last_Row As Double
  Dim Summary_Table_Row As Integer

  Dim Greatest_Percentage_Increase As Double
  Dim Greatest_Percentage_Decrease As Double
  Dim Greatest_Total_Volume As Double
 
    ' Begins loop which will run through each worksheet in the workbook.
    For Each ws In ThisWorkbook.Sheets

      ' Populates the headings required for each worksheet.
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change ($)"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"

      ' Formats the above heading's interior and font to be clearly distinguishable from the raw data.   
      ws.Range("I1:L1,O1:O4,P1:Q1").Font.Color = RGB(255,255,255)
      ws.Range("I1:L1,O1:O4,P1:Q1").Interior.Color = RGB(0,0,0)
    
    
      ' Defining the variables for the currently observed worksheet:
      
      Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row                    ' Defines the last populated row in column one of the current sheet.
      Summary_Table_Row = 2                                               ' Defines the initial row that each unqiue ticker and its' yearly change, percent change, and total stock volume are inserted into.
      Total_Stock_Volume = 0                                              ' Assigns the total stock volume variable to a initial value of 0 to the current sheet observed.
      Opening_Price = ws.Cells(2, 3).Value                                ' Defines the intial value of the opening price.

        ' Begins a loop to recover the values needed for our summary table.
        For i = 2 To Last_Row
        
        Closing_Price = ws.Cells(i, 6).Value                              ' Defines closing price as the current row's - column 6 value, has to be done within the loop as reliant on i's value. 
                
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then        ' Conditional: checks if the vertically neigbouring ticker symbols observed in column one differ within the set range.
        
          Ticker_Symbol = ws.Cells(i, 1).Value                            ' When condition is met, the uppermost obervation is stored within the Ticker_Symbol variable.
          Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value              ' When condition is met, the total volume in the current row's column is added onto the value currently stored in Total Volume (0).
          Yearly_Change = Closing_Price - Opening_Price                   ' When condition is met, the opening_price is deducted from the closing price and stored within the Yearly_Change variable.
     
          ' When condition is met, displays the value's stored in the Ticker_Symbol, Yearly_Change, and Total_Stock_Volume Variables into the appropritate location in the summary table.
          ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
          ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
          ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
          

              'Nested loop helps apply conditional formatting to negative and positive yearly change values as they are recorded in the summary table.
              
              If ws.Range("J" & Summary_Table_Row) > 0 Then                     ' Conditional: If the observed value in the summary table row/ yearly change column is positive...
              ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0,255,0)   ' Sets that yearly change cell's interior to the colour green. RGB used as a safeguard in case the colour  palette indexes change in future.

              ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then           ' Conditional: If the observed value in the summary table row/ yearly change column is negative...
              ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255,0,0)   ' Sets that yearly change cell's interior to the colour red.
              
              End If
         
          ' Percent change variable is assigned to equal the (closing-opening)/ opening. Where (closing - opening) is the yearly change already computed. 
          Percent_Change = Yearly_Change / Opening_Price
          ' When condition is met, the percent change is displayed in the summary table and formatted as a percentage number format. 
          ws.Range("K" & Summary_Table_Row).Value = Percent_Change
          ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"      

               'Nested loop helps apply conditional formatting to negative and positive percentage change values as they are recorded in the summary table.
              
              If ws.Range("K" & Summary_Table_Row) > 0 Then                     ' Conditional: If the observed value in the summary table row/ percentage change column is positive...
              ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(0,255,0)   ' Sets that yearly change cell's interior to the colour green. RGB used as a safeguard in case the colour  palette indexes change in future.

              ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then           ' Conditional: If the observed value in the summary table row/ percentage change column is negative...
              ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(255,0,0)   ' Sets that yearly change cell's interior to the colour red.
              
              End If

          ' Adds one onto the summary table row so next time the condition is met in the loop, the values can be displayed in the correct position.
          Summary_Table_Row = Summary_Table_Row + 1

          ' When condition is met, after i = 2 is met in the loop, the opening price will always be in the row below the one being observed. The below helps define this. 
          Opening_Price = ws.Cells(i + 1, 3).Value
          
          ' When condition is met, resets Total Stock Volume to zero as a new ticket symbol will be observed on the next loop. 
          Total_Stock_Volume = 0
        
        
          ' For any instances where the condition is not met, the tickers are of the same type and therefore, the row's volume will need to be added onto the total stored in Total_Volume. 
          Else
          Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
     
          End If
              
     
        Next i 
      
        ' Adding the functionalitiy to this script to return the greatest percentage increase and total volume, as well as the greatest percentage decrease from our summary table. 
        
        Greatest_Percentage_Increase = WorksheetFunction.max(ws.range("K:K"))       ' Extracts the highest value from column K and assigns it to the variable 'Greatest_Percentage_Increase"
        ws.Range("Q2").Value = Greatest_Percentage_Increase                         ' Displays the greatest percentage increase in a new summary table.  
        ws.Range("Q2").NumberFormat = "0.00%"                                       ' Converts the number format of the displayed greatest percentage increase into a percentage. 

        Greatest_Percentage_Decrease = WorksheetFunction.min(ws.range("K:K"))       ' Extracts the lowest value from column K and assigns it to the variable 'Greatest_Percentage_Decrease"
        ws.Range("Q3").Value = Greatest_Percentage_Decrease                         ' Displays the greatest percentage increase in the new summary table.  
        ws.Range("Q3").NumberFormat = "0.00%"                                       ' Converts the number format of the displayed greatest percentage decrease into a percentage. 

        Greatest_Total_Volume = WorksheetFunction.max(ws.range("L:L"))              ' Extracts the highest value from column L and assigns it to the variable 'Greatest_Total_Volume"
        ws.Range("Q4").Value = Greatest_Total_Volume                                ' Displays the greatest total volume in the new summary table.  

        ' Now the ticker symbols linked to the values in the new summary table need to be extracted. 

        ' Redefines the last row to be the last populated row in column 11. 
        Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row

          ' Begins a loop to search through each of the percentage change and total stock volume columns in our summary table. 
          For i = 2 To Last_Row
            For j = 11 To 12

              If ws.Cells(i, j).Value = Greatest_Percentage_Increase Then           ' Conditional: If the observed cell matches with the greatest percentage increase...
              ws.Range("P2").Value = ws.Cells(i, 9).Value                           ' When condition is met, the corresponding ticker symbol is stored in the new summary table. 

              ElseIf ws.Cells(i,j).Value = Greatest_Percentage_Decrease Then        ' Conditional: If the observed cell matches with the greatest percentage decrease...
              ws.Range("P3").Value = ws.Cells(i, 9).Value                           ' When condition is met, the corresponding ticker symbol is stored in the new summary table. 
    
              ElseIf ws.Cells(i, j).Value = Greatest_Total_Volume Then              ' Conditional: If the observed cell matches with the greatest total volume...
              ws.Range("P4").Value = ws.Cells(i, 9).Value                           ' When condition is met, the corresponding ticker symbol is stored in the new summary table. 
    
              End If
    
            Next j

          Next i

        ' Autofits the worksheet's columns to make all data visable and clear.
        ws.Cells.EntireColumn.Autofit
  
  Next ws
     
End Sub





