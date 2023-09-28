Sub MultipleYearStockData():

    For Each ws In Worksheets
        'Assign variables
        Dim WorksheetName As String
       'Pull the WorksheetName
        WorksheetName = ws.Name
       
        'Current row
         Dim LastRowA As Long
        'last row column I
        Dim LastRowI As Long
       
        'Variable for % change calculation
        Dim PerCnge As Double
        'Variable for greatest increase calculation
        Dim GreatIncrease As Double
        'Variable for greatest decrease calculation
        Dim GreatDecrease As Double
        'Variable for greatest total volume
        Dim GreatVoume As Double
        
        Dim TickCount As Long
        'Last row column A
        
        Dim i As Long
        'Start row of ticker
        Dim j As Long
        'Index counter to fill Ticker row
        
        
        'Define column for each workseet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter to start at first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)
        
            'Setup loop to go through all rows
            For i = 2 To LastRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I 
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and write Yearly Change in column J 
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                'Calculate and write percent change in column K 
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Fromatting cells for % change
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                
                    'Conditional formating to assign colors based on % change +/-
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
                    'Set cell background color to red if negative % change in ticker
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green if positive positive % change in ticker
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                                        
                'Calculate and write total volume in column L 
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row of the ticker 
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
        'Assign range to cells for summary data
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Summary loop 
            
                For i = 2 To LastRowI
               'Find the greatest % increase and populate the cell
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
            
               'Find the greatest % decrease and poulate the cell
                If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If

                'Find the greatest volume and populate the cell
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
                             
             'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next i
            
        'Adjust column width 
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
