Sub group_by_ticker():

    'setting the variables
    Dim n, i, Yearly_change, Percent_change As Integer
    Dim ticker As String
    Dim total_stock_volume, open_value, close_value As Double
    Dim sumary_table_row As Integer
    Dim percent_change1 As Double
    Dim greatest_ticker, lowest_ticker, greatest_volume_ticker As String
    Dim greatest, lowest, greatest_volume As Double
    Dim A, WS, y As Integer
    Dim current As Worksheet

    WS = ActiveWorkbook.Worksheets.Count
    y = 2018

    'looping between worksheets
    For Each current In Worksheets:
    
        'Counting how many rows are for looping
        n = current.UsedRange.Rows.Count
    
        'inicialize values
    
        sumary_table_row = 2
        total_stock_volume = 0
        Yearly_change = 0
        greatest = 0
        lowest = 0
        greatest_volume = 0
         
        'starting the loop
        For i = 2 To n:
            'saving the opening value
            If Cells(i - 1, 1).Value <> Cells(i, 1) Then
                open_value = Cells(i, 3).Value
                
            End If
            
           'if is the last same value of the row then:
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
                close_value = Cells(i, 6).Value
                ticker = Cells(i, 1).Value
                total_stock_volume = total_stock_volume + Cells(i, 7)
                Yearly_change = close_value - open_value
                percent_change1 = (close_value - open_value) / open_value
        
            
                'printing the values on the table
                Range("i" & sumary_table_row).Value = ticker
                Range("L" & sumary_table_row).Value = total_stock_volume
                Range("j" & sumary_table_row).Value = Yearly_change
                Range("K" & sumary_table_row).Value = percent_change1
                Range("M" & sumary_table_row).Value = open_value
                Range("N" & sumary_table_row).Value = close_value
                'conditional formating
                If Yearly_change < 0 Then
                Range("j" & sumary_table_row).Interior.ColorIndex = 3
                Else:
                Range("j" & sumary_table_row).Interior.ColorIndex = 4
                End If
                'percentage format
                Range("K" & sumary_table_row).Style = "percent"
                Range("R2").Style = "percent"
                Range("R3").Style = "percent"
                
                  'greatest and lowest percentage rate
                If percent_change1 > greatest Then
                    greatest = percent_change1
                    Range("R2").Value = greatest
                    Range("Q2").Value = ticker
                End If
                
                 If percent_change1 < lowest Then
                    lowest = percent_change1
                    Range("R3").Value = lowest
                    Range("Q3").Value = ticker
                End If
                
                 If total_stock_volume > greatest_volume Then
                    greatest_volume = total_stock_volume
                    Range("R4").Value = greatest_volume
                    Range("Q4").Value = ticker
                End If
                
                
                'reset of the values
                sumary_table_row = sumary_table_row + 1
                total_stock_volume = 0
                Yearly_change = 0
                
              
                    
                
                
            Else:
                total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
            'end of if
            End If
            
        Next i
        y = y + 1
    Next
End Sub