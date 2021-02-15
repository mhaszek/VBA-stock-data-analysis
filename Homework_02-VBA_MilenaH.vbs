Attribute VB_Name = "Module2"
'Homework for Data Analytics BootCamp - Milena Hablaszek
Sub VBAchallenge():

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row for the main dataset
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).row + 1
        
        ' Add Headers to Summary Table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Define variables
        Dim row As Integer
        row = 2
        
        Dim rows_count As Integer
        rows_count = 0
  
        ' total had to be set to 'decimal' data type, because 'long' was returing overflow error
        Dim total As Double
        total = CDec(total)
        total = 0
        
        Dim opening As Double
        Dim closing As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        
                        
        ' Loop through all rows in the dataset except Headers
        For i = 2 To lastRow
          
            ' Search for when the value of the next cell is the same than that of the current cell
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
                ' Add value from Column G for each row with the same ticker
                total = total + ws.Cells(i, 7).Value
                
                ' Increase count of rows with the same ticker
                rows_count = rows_count + 1
                
            ' Search for when the value of the next cell is different than that of the current cell
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Asing value of Cell(current row 'i', Column A) to Cell(row 'row', Column I)
                ws.Cells(row, 9) = ws.Cells(i, 1).Value
         
                ' Add value from Column G for current row 'i' to the total of the same ticker
                total = total + ws.Cells(i, 7).Value
      
                ' Asing value of 'total' for current ticker to Cell(row 'row', Column L)
                ws.Cells(row, 12) = total
      
                ' Find closing price value for current ticker Cell(current row 'i', Column F) = last record for this ticker
                closing = ws.Cells(i, 6).Value
                                
                ' Find opening price value for current ticker Cell(row 'i - rows_count', Column C) = first record for this ticker
                opening = ws.Cells((i - rows_count), 3).Value
                               
                ' Calculate Yearly Change from opening price to the closing price
                yearly_change = closing - opening
                
                ' Asing value of 'yearly_change' for current ticker to Cell(row 'row', Column J)
                ws.Cells(row, 10) = yearly_change
                                
                ' Change Number Formatting for Cell(row 'row', Column J) to look nicer
                ws.Cells(row, 10).NumberFormat = "0.00"
                
                ' Calculate Percent Change from opening price to the closing price. IF() function used to avoid error when diving by "0"
                If (opening = 0 Or closing = 0) Then
                
                    percent_change = 0
                
                Else
                    
                    percent_change = (closing - opening) / opening
                
                End If
                                
                ' Asing value of 'percent_change' for current ticker to Cell(row 'row', Column K)
                ws.Cells(row, 11) = Round(percent_change, 4)
                
                ' Change Number Formatting for Cell(row 'row', Column K) to look nicer
                ws.Cells(row, 11).NumberFormat = "0.00%"
                
                ' Change Color Formatting for Cell(row 'row', Column J) to highlight positive change (>0) in green and negative change (<0) in red
                If (yearly_change > 0) Then
                
                    ws.Cells(row, 10).Interior.ColorIndex = 4

                ElseIf (yearly_change < 0) Then
                
                    ws.Cells(row, 10).Interior.ColorIndex = 3
                
                End If
                
                ' Increase count of row so that the data for next ticker will be added in a new row in the Summary Table
                row = row + 1
      
                ' Reset count of rows with the same ticker so that the counter will start again for the next ticker
                rows_count = 0
                
                ' Reset total for the same ticker so that the sum will start again for the next ticker
                total = 0
                
            End If

        Next i

'-----------------------------------------------------------------------------------------------------------------------------
' ### CHALLENGES
' Challenges have been added to the same file and same loop so that the script can be run only once for all sheets and return all required outcomes.
        
        ' Find the last row for the Summary Table
        lastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).row
        
        ' Add Headers to the Challenge Table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
      
        ' Find Max value of Percent Change in Column K and assign it to Cell(Q2)
        increase = WorksheetFunction.Max(ws.Range("K2:K" & lastRow2))
        ws.Range("Q2").Value = increase

        ' Find the Ticker corresponding with the Greatest Increase in Column I and assign it to Cell(P2)
        match_increase = WorksheetFunction.Match(increase, ws.Range("K2:K" & lastRow2), 0)
        ws.Range("P2").Value = ws.Cells(match_increase + 1, 9)

        ' Find Min value of Percent Change in Column K and assign it to Cell(Q3)
        decrease = WorksheetFunction.Min(ws.Range("K2:K" & lastRow2))
        ws.Range("Q3").Value = decrease

        ' Find the Ticker corresponding with the Greatest Decrease in Column I and assign it to Cell(P3)
        match_decrease = WorksheetFunction.Match(decrease, ws.Range("K2:K" & lastRow2), 0)
        ws.Range("P3").Value = ws.Cells(match_decrease + 1, 9)
        
        ' Find Max value of Total Stock Volume in Column L and assign it to Cell(Q4)
        total_value = WorksheetFunction.Max(ws.Range("L2:L" & lastRow2))
        ws.Range("Q4").Value = total_value

        ' Find the Ticker corresponding with the Greatest Total Value in Column I and assign it to Cell(P4)
        match_total = WorksheetFunction.Match(total_value, ws.Range("L2:L" & lastRow2), 0)
        ws.Range("P4").Value = ws.Cells(match_total + 1, 9)
        
        ' Change Formatting so that results look nice
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("I:Q").AutoFit
        
    Next ws
  

End Sub

