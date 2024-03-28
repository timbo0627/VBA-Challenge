Sub stock()
    Dim stock As String
    Dim startprice As Double
    Dim endprice As Double
    Dim change As Double
    Dim perchange As String
    Dim tablecount As Long
    Dim volume As LongLong
    
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    Range("A1").Select
    For Each ws In Worksheets
        Dim x As Long
        x = 0
        worksheetname = ws.Name
        tablecount = 2
        For x = 2 To NumRows
            If x = 2 Then
                startprice = ws.Cells(x, 3).Value
                stock = ws.Cells(x, 1).Value
                volume = ws.Cells(x, 7).Value
                endprice = ws.Cells(x, 6).Value
                ws.Cells(1, 10).Value = "Ticker"
                ws.Cells(1, 11).Value = "Yearly Change"
                ws.Cells(1, 12).Value = "Percent Change"
                ws.Cells(1, 13).Value = "Total Stock Volume"
            ElseIf ws.Cells(x, 1).Value = stock Then
                volume = volume + ws.Cells(x, 7).Value
                endprice = ws.Cells(x, 6).Value
            Else
                change = endprice - startprice ' set change to final price minus start price
                perchange = change / startprice ' set perchange as change divided by start price
                ws.Cells(tablecount, 10).Value = stock
                ws.Cells(tablecount, 11).Value = change
                ws.Cells(tablecount, 12).Value = perchange
                ws.Cells(tablecount, 12).NumberFormat = "0.00%"
                ws.Cells(tablecount, 13).Value = volume
                ' format cell colors based on yearly change
                If ws.Cells(tablecount, 11).Value >= 0 Then
                    ws.Cells(tablecount, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(tablecount, 11).Interior.ColorIndex = 3
                End If
                ' reset variables
                stock = ws.Cells(x, 1).Value
                startprice = ws.Cells(x, 3).Value
                endprice = ws.Cells(x, 6).Value
                tablecount = tablecount + 1
                volume = ws.Cells(x, 7).Value
            End If
        Next
        ' add best worst and volume to each sheet
        Dim bestie As String
        Dim worstie As String
        Dim bestvol As String
        Dim bestieval As Double
        Dim worstieval As Double
        Dim bestvolval As LongLong
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        For y = 2 To NumRows
            If ws.Cells(y, 12).Value > bestieval Then
                bestie = ws.Cells(y, 10).Value
                bestieval = ws.Cells(y, 12).Value
            End If
            If ws.Cells(y, 12).Value < worstieval Then
                worstie = ws.Cells(y, 10).Value
                worstieval = ws.Cells(y, 12).Value
            End If
            If ws.Cells(y, 13).Value > bestvolval Then
                bestvol = ws.Cells(y, 10).Value
                bestvolval = ws.Cells(y, 13).Value
            End If
        Next
        ws.Cells(2, 17).Value = bestie
        ws.Cells(2, 18).Value = bestieval
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = worstie
        ws.Cells(3, 18).Value = worstieval
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = bestvol
        ws.Cells(4, 18).Value = bestvolval
    
        ws.Range("A1:R1").Columns.AutoFit
        
    
    Next ws
End Sub
