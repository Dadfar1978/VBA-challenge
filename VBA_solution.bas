Attribute VB_Name = "Module1"
Sub stockanalysis():
    
    Dim ws As Worksheet
    Dim i As Long
    Dim Lastrow As Long
    Dim Ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim changePrice As Double
    Dim percentChange As Double
    Dim Total As Double
    Dim Row As Integer
    Dim MaxpercentChange As Double
    Dim MinpercentChange As Double
    Dim MaxTotalVolume As Double
    Dim MaxpercentTicker As String
    Dim MinpercentTicker As String
    Dim MaxVolumeTicker As String
    
    For Each ws In Worksheets
    
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarterly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Total = 0
        Row = 2
        openingPrice = ws.Range("C" & 2).Value
        
        
     For i = 2 To Lastrow
        
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                 closingPrice = ws.Range("F" & i).Value
                 changePrice = closingPrice - openingPrice
                     
                     If ws.Range("I" & Row).Value > 0 Then
                        ws.Range("I" & Row).Interior.ColorIndex = 4
                    ElseIf ws.Range("I" & Row).Value < 0 Then
                        ws.Range("I" & Row).Interior.ColorIndex = 3
                    Else
                    End If
                 
                If openingPrice <> 0 Then
                   percentChange = changePrice / openingPrice
                 
                Else
                 percentChange = 0
                End If
                 ws.Range("H" & Row).Value = Ticker
                 ws.Range("I" & Row).Value = changePrice
                 ws.Range("J" & Row).Value = percentChange
                 ws.Range("J" & Row).NumberFormat = "0.00%"
                 ws.Range("K" & Row).Value = Total
                Row = Row + 1
                Total = 0
              If (i + 1) <= Lastrow Then
                openingPrice = ws.Cells(i + 1, 3).Value
             End If
                
           Else
           
           Total = Total + ws.Cells(i, 7).Value
              
    End If
            
            
            
     Next i
     MaxpercentChange = WorksheetFunction.Max(ws.Range("J2:J" & Row))
     MinpercentChange = WorksheetFunction.Min(ws.Range("J2:J" & Row))
     MaxTotalVolume = WorksheetFunction.Max(ws.Range("K2:K" & Row))
     MaxPercentRow = WorksheetFunction.Match(MaxpercentChange, ws.Range("J2:J" & Row - 1), 0) + 1
     MaxpercentTicker = ws.Range("H" & MaxPercentRow).Value
     MinPercentRow = WorksheetFunction.Match(MinpercentChange, ws.Range("J2:J" & Row - 1), 0) + 1
     MinpercentTicker = ws.Range("H" & MinPercentRow).Value
     MaxVolumeRow = WorksheetFunction.Match(MaxTotalVolume, ws.Range("K2:K" & Row - 1), 0) + 1
     MaxVolumeTicker = ws.Range("H" & MaxVolumeRow).Value
     
     
     ws.Range("P2").Value = MaxpercentChange
     ws.Range("P3").Value = MinpercentChange
     ws.Range("P2:P3").NumberFormat = "0.00%"
     ws.Range("P4").Value = MaxTotalVolume
     ws.Range("O2").Value = MaxpercentTicker
     ws.Range("O3").Value = MinpercentTicker
     ws.Range("O4").Value = MaxVolumeTicker
     
     
     
    Next ws
End Sub
