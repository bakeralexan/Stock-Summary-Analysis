Sub AlphabeticalTesting():


        Dim header As Boolean
        header = false

        Dim ws As Worksheet
        
        For Each ws in Worksheets
            Dim EndRow As Long
            Dim SumRow As Long
            Dim YearlyChange As Double
            Dim TotalStockVolume As Double
            Dim PercentChange As Double
            Dim OpenPrice As Double
            Dim ClosePrice As Double
            Dim i As Long
            Dim TickerMax As String
            Dim TickerMin As String
            Dim TickerVol As String

            Dim GreatestIncrease As Double
            Dim GreatestDecrease As Double
            Dim GreatestVolume As Double
            GreatestIncrease = 0
            GreatestDecrease = 0
            GreatestVolume = 0
            TickerMax = ""
            TickerMin = ""
            TickerVol = ""


            EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            SumRow = 2
            YearlyChange = 0
            TotalStockVolume = 0
            PercentChange = 0
            OpenPrice = 0
            ClosePrice = 0

            If header = false then
                ws.Range("I1").Value = "Ticker" 
                ws.Range("J1").Value = "Yearly Change" 
                ws.Range("K1").Value = "Percent Change" 
                ws.Range("L1").Value = "Total Stock Volume"
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("P1").Value = "Ticker"
                ws.Range("Q1").Value = "Value"
            Else 
                header = true
            End If



            OpenPrice = ws.Cells(2, 3).Value 
            
            For i = 2 To EndRow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    Ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & SumRow).Value = Ticker
                    ClosePrice = ws.Cells(i, 6).Value 
                    YearlyChange = ClosePrice - OpenPrice
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                    ws.Range("L" & SumRow).Value = TotalStockVolume
                    ws.Range("J" & SumRow).Value = YearlyChange
                    If OpenPrice <> 0 Then
                        PercentChange = (YearlyChange / OpenPrice)
                    Else
                        PercentChange = 0
                    End If

                    ws.Range("K" & SumRow).Value = PercentChange
                    ws.Range("K" & SumRow).Style = "Percent"

                    If (YearlyChange > 0) Then  
                        ws.Range("J" & SumRow).Interior.ColorIndex = 4 
                    ElseIf (YearlyChange <= 0) Then
                        ws.Range("J" & SumRow).Interior.ColorIndex = 3
                    End If

                    If PercentChange > GreatestIncrease Then
                        GreatestIncrease = PercentChange
                        TickerMax = Ticker
                    Elseif PercentChange < GreatestDecrease Then
                        GreatestDecrease = PercentChange
                        TickerMin = Ticker
                    End if
                    If TotalStockVolume > GreatestVolume Then
                        GreatestVolume = TotalStockVolume
                        TickerVol = Ticker
                    End If

                    SumRow = SumRow + 1

                    OpenPrice = ws.Cells(i + 1, 3).Value

                    TotalStockVolume = 0
                    
                    
                Else   
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                    
                End If


            Next i
            ws.Range("Q2").Value = GreatestIncrease
            ws.Range("Q3").Value = GreatestDecrease
            ws.Range("Q4").Value = GreatestVolume
            ws.Range("P3").Value = TickerMin
            ws.Range("P2").Value = TickerMax
            ws.Range("P4").Value = TickerVol
            ws.Range("Q2:Q3").Style = "Percent"
            
        Next ws


        

    MsgBox ("Multiple Sheets' Stocks Summarized")
End Sub