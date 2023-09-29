Attribute VB_Name = "Module1"
' first loop through each sheet - 'For Each' reference all 'Worksheets'

Sub StockPart1():

    For Each ws In Worksheets
 
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker block
        Dim j As Long
        'Index counter to fill Ticker row
        Dim Ticker As Long
        'Last row column A
        Dim LastRowA As Long
        'last row column I
        Dim LastRowI As Long
        'Variable for percent change calculation
        Dim PerChange As Double
        'Variable for greatest increase calculation
        Dim GreatIncr As Double
        'Variable for greatest decrease calculation
        Dim GreatDecr As Double
        'Variable for greatest total volume
        Dim GreatVol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker to first row
        Ticker = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
        
            'Loop through all rows
            For i = 2 To LastRow
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I (#9)
                ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and write Yearly Change in column J (#10)
                ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating
                    If ws.Cells(Ticker, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(Ticker, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(Ticker, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in column L (#12)
                ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase Ticker by 1
                Ticker = Ticker + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
        'Prepare for summary
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'For greatest total volume
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
                'For greatest increase
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
                'For greatest decrease
                If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
         Next i
            
    
            
    Next ws
End Sub

