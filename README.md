# VBA-challenge
2nd homework assignment
Included are pictures of the files with completed each of the tabs completed using the code.
Alphabetical Testing with VBA code (instead of multiple_year_stock_data excel sheet since it's too large).

****Received help from Tyler Aden And also worked with Brandon Britt to clean up code and ensure that the code ran for every tab consecutively.

Below is the code for the homework which is also included in the Alphabetical Testing excel sheet(as mentioned in line 4 above)

Sub stock_analysis()
    
    Dim i As Long
    Dim j As Integer
    Dim total As Double
    Dim rowCount As Long
    Dim percentChange As Double
    Dim change As Double
    Dim num_max As Double
    Dim num_min As Double
    Dim num_vol As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets
                
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    Start = 2
    total = 0
    change = 0
    j = 0
    
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To rowCount
    
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            
            total = total + ws.Cells(i, 7).Value
            
            If total = 0 Then
                
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            
            Else
            
                If ws.Cells(Start, 3) = 0 Then
                    For find_val = Start To i
                        If ws.Cells(find_val, 3) <> 0 Then
                            Start = find_val
                            Exit For
                        End If
                    Next find_val
                End If
             
                change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
                percentChange = change / ws.Cells(Start, 3)
             
                Start = i + 1
             
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = change
                ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = total
                
                If change > 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                ElseIf change < 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End If
            End If
                
                total = 0
                change = 0
                j = j + 1
            
            Else
                total = total + ws.Cells(i, 7).Value
                
        End If
    
    
    Next i
    
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
    
    ws.Range("P2").Value = ws.Range("I" & increase_number + 1).Value
    ws.Range("P3").Value = ws.Range("I" & decrease_number + 1).Value
    ws.Range("P4").Value = ws.Range("I" & volume_number + 1).Value
    
    Next ws
    
End Sub

