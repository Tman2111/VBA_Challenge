Sub Multiple_Sheets()
    Dim xSh As worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
' New Column Headers
    Cells(1, 9).Value = "Ticker Symbol"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Find last active row
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    
' Variables for Ticker Symbol to Summary Loop
    Dim TickerSymbol As String
    Dim TotalVolume As Double
    Dim SummaryRow As Integer
    Dim YearlyChange As Double
    Dim PercentChange As Double
        TotalVolume = 0
        SummaryRow = 0
        begin = 2
        
'Loop for Ticker Symbol and Volume to Summary

    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            TotalVolume = TotalVolume + Cells(i, 7).Value
            TickerSymbol = Cells(i, 1).Value
            
            If TotalVolume = 0 Then
                
            
                Range("I" & 2 + SummaryRow).Value = TickerSymbol
                Range("J" & 2 + SummaryRow).Value = 0
                Range("K" & 2 + SummaryRow).Value = "%" & 0
                Range("L" & 2 + SummaryRow).Value = 0
            Else
                If Cells(begin, 3) = 0 Then
                    For k = begin To i
                       If Cells(k, 3).Value <> 0 Then
                            begin = k
                            Exit For
                        End If
                            
                    Next k
                End If
                
                'Calculating the yearly change and percent change
                YearlyChange = (Cells(i, 6) - Cells(begin, 3))
                PercentChange = Round((YearlyChange / Cells(begin, 3) * 100), 2)
                
                begin = i + 1
                'print our results
                    
                    Range("I" & 2 + SummaryRow).Value = TickerSymbol
                    Range("J" & 2 + SummaryRow).Value = Round(YearlyChange, 2)
                    Range("K" & 2 + SummaryRow).Value = "%" & PercentChange
                    Range("L" & 2 + SummaryRow).Value = TotalVolume
                    
                
                    
            'reset the variables for a new ticker symbol
            End If
            
            YearlyChange = 0
            TotalVolume = 0
            SummaryRow = SummaryRow + 1
            
    
    Else
    
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
    End If

Next i
        
        
'Conditional format yearl
dataRowStart = 2
    dataRowEnd = Cells(Rows.Count, "I").End(xlUp).Row
    For i = dataRowStart To dataRowEnd
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = vbGreen
        Else
            Cells(i, 10).Interior.Color = vbRed
        End If
    Next i
        
    For i = dataRowStart To dataRowEnd
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.Color = vbGreen
        Else
            Cells(i, 11).Interior.Color = vbRed
        End If
    Next i
    
    
End Sub