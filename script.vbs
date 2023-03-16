Sub Multiple_Year_Stock()

'Defining all  variables
Dim opening, closing, change, total_stock_volme, percent_change, count, num As Integer
Dim ws As Worksheet

'To function in every worksheet
For Each ws In Worksheets

'Assigning title
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Assigning value
count = 1
num = 2
total_stock_volme = 0
    
'To count the last row of coumn A
    
lastrow = ws.Cells(Rows.count, "A").End(xlUp).Row

'For Ticker symbol, yearly change, percent change, and total stock volume
For i = 2 To lastrow
            
        'If Ticker symbol is not equal then if condition is true
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(num, 9).Value = ws.Cells(i, 1).Value
        count = count + 1
            
        'Assigning the closing value and opening value
        opening = ws.Cells(count, 3).Value
        closing = ws.Cells(i, 6).Value
            
                ' Adding total stock volume for each ticker symbol
                For j = count To i
                total_stock_volme = total_stock_volme + ws.Cells(j, 7).Value
                Next j
            
        'For change
            
        If opening = 0 Then
        percent_change = closing
        Else
        change = closing - opening
        percent_change = change / opening
        End If

        'Keeping value of cells
        ws.Cells(num, 10).Value = change
        ws.Cells(num, 11).Value = percent_change
        ws.Cells(num, 11).NumberFormat = "0.00%"
        ws.Cells(num, 12).Value = total_stock_volme
              
        num = num + 1
            
        'Resetting value for each ticker
        total_stock_volme = 0
        change = 0
        percent_change = 0
            
        'For not repeating the value of count
        count = i
        
        End If
Next i
    
'-----------------------------------------------------------------


'For Greatest % increase, Greatest % decrease, and Greatest total volume
    
'For the last row of column K
last = ws.Cells(Rows.count, "K").End(xlUp).Row
    
'Defining variable
Increase = 0
Decrease = 0
greatest = 0
    
For k = 3 To last
m = k - 1
                        
        'Defining the cells for increase, decrease and volume
        new_k = ws.Cells(k, 11).Value
        old_k = ws.Cells(m, 11).Value
        new_volume = ws.Cells(k, 12).Value
        old_volume = ws.Cells(m, 12).Value
            
            
        'Finding greatest increase
        If Increase > new_k And Increase > old_k Then
        Increase = Increase
        ElseIf new_k > Increase And new_k > old_k Then
        Increase = new_k
        increase_name = ws.Cells(k, 9).Value
        ElseIf old_k > Increase And old_k > new_k Then
        Increase = old_k
        increase_name = ws.Cells(m, 9).Value
        End If
     
        'Finding greatest decrease
        If Decrease < new_k And Decrease < old_k Then
        Decrease = Decrease
        ElseIf new_k < Increase And new_k < old_k Then
        Decrease = new_k
        decrease_name = ws.Cells(k, 9).Value
        ElseIf old_k < Increase And old_k < new_k Then
        Decrease = old_k
        decrease_name = ws.Cells(m, 9).Value
        End If
  
        'Finding the greatest volume
        If greatest > new_volume And greatest > old_volume Then
        greatest = greatest
        ElseIf new_volume > greatest And new_volume > old_volume Then
        greatest = new_volume
        ElseIf old_volume > greatest And old_volume > new_volume Then
        greatest = old_volume
        greatest_name = ws.Cells(m, 9).Value
        End If
            
Next k
  
' Assigning names and getting value for greatest increase,greatest decrease, and  greatest volume
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker Name"
ws.Range("P1").Value = "Value"
ws.Range("O2").Value = increase_name
ws.Range("O3").Value = decrease_name
ws.Range("O4").Value = greatest_name
ws.Range("P2").Value = Increase
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").Value = Decrease
ws.Range("P3").NumberFormat = "0.00%"
ws.Range("P4").Value = greatest
    
    
'--------------------------------------------------

' Conditional formatting thathighlights positive change in green and negative change in red

finish = ws.Cells(Rows.count, "J").End(xlUp).Row
For j = 2 To finish
        If ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
Next j
    
'For every Worksheet
Next ws

End Sub

