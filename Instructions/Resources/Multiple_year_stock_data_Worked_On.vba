Sub vbaHW()

'Loop through all the sheets
For i = 1 To ActiveWorkbook.Worksheets.Count
'For i = 3 To 3 'for testing purposes, the above line is the real one
        With ActiveWorkbook.Worksheets(i)
        '==========Begin per sheet code (and initialize variables)==========
        Dim pivotCounter As Integer: pivotCounter = 2  'Count the row of the pivot table we're writting to
        Dim ticker As String 'Save current string
        Dim totalVol As Double: totalVol = 0  'Count total volume
        Dim yearOpen As Double: yearOpen = 0
        Dim yearClose As Double: yearClose = 0
        Dim yearDiff As Double
        'Write pivot headers
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
        
        
        'Loop through all the rows of each sheet
        yearOpen = .Range("C2").Value
        For r = 2 To .Range("A" & .Rows.Count).End(xlUp).Row
            
            'Read all data of row
            ticker = .Range("A" & r).Value
            totalVol = totalVol + .Range("G" & r).Value

            
            If .Range("A" & (r + 1)).Value <> .Range("A" & r).Value Then
                'Write ticker
                .Range("I" & pivotCounter).Value = ticker
                
                'Opening and closing prices and their reset
                yearClose = .Range("F" & r).Value
                
                'Yearly Change stuff
                yearDiff = yearClose - yearOpen
                .Range("J" & pivotCounter).Value = yearDiff
                If yearDiff > 0 Then .Range("J" & pivotCounter).Interior.ColorIndex = 4
                If yearDiff < 0 Then .Range("J" & pivotCounter).Interior.ColorIndex = 3
                
                'Percent Change stuff
                If (yearOpen = 0) Or (yearClose = 0) Then
                    .Range("K" & pivotCounter).Value = 0
                Else
                    .Range("K" & pivotCounter).Value = yearDiff / yearOpen
                End If
                .Range("K" & pivotCounter).NumberFormat = "0.00%"
                
               
                
                'Reset opening and closing (year closing is assigned above so year Open needs to be reset)
                yearOpen = .Range("F" & (r + 1)).Value
            
                'Write and reset total volume
                .Range("L" & pivotCounter).Value = totalVol
                totalVol = 0
                
                'Pivot row done.
                pivotCounter = pivotCounter + 1
            End If
        Next r
        
        'Bonus Pivot
        .Range("O2").Value = "Greatest % Increase": .Range("O3").Value = "Greatest % Decrease": .Range("O4").Value = "Greatest Total Volume"
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        .Range("Q2:Q3").NumberFormat = "0.00%"
        Dim topInc As Double, topDec As Double, topVol As Double 'In hindsight you probably don't need this
        'Write down the first row/initialize it all. Thing below is the ticker
        .Range("P2:P4").Value = .Range("I2").Value
        'Greatest % increase
        .Range("Q2").Value = .Range("K2").Value: topInc = .Range("K2").Value
        'Greatest % decrease
        .Range("Q3").Value = .Range("K2").Value: topDec = .Range("K2").Value
        'Greatest vol
        .Range("Q4").Value = .Range("L2").Value: topVol = .Range("L2").Value
        
        For b = 3 To .Range("I" & .Rows.Count).End(xlUp).Row
            'Check top increase
            If .Range("K" & b) > topInc Then
                .Range("P2").Value = .Range("I" & b).Value 'Ticker
                .Range("Q2").Value = .Range("K" & b).Value 'Value
                topInc = .Range("K" & b).Value 'New Top
            End If
            'Check top decrease
            If .Range("K" & b) < topDec Then
                .Range("P3").Value = .Range("I" & b).Value 'Ticker
                .Range("Q3").Value = .Range("K" & b).Value 'Value
                topDec = .Range("K" & b).Value 'New Bot
            End If
            'Check top volume
            If .Range("L" & b) > topVol Then
                .Range("P4").Value = .Range("I" & b).Value 'Ticker
                .Range("Q4").Value = .Range("L" & b).Value 'Value
                topVol = .Range("L" & b).Value 'New Top
            End If
        Next b
        
'END looping through all the sheets
        End With
Next i
        
    

End Sub

