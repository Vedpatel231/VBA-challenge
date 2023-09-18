Attribute VB_Name = "Module1"
Sub AnalyzeStocks()

    
    Dim LR As Long
    Dim counter As Long
    Dim Ticker As String
    Dim OpenPrice As Double, ClosePrice As Double
    Dim YearlyChange As Double, PercentChange As Double
    Dim TotalVolume As Double
    Dim GreatestPercentIncrease As Double, GreatestPercentDecrease As Double, GreatestTotalVolume As Double
    Dim GreatestPercentIncreaseTicker As String, GreatestPercentDecreaseTicker As String, GreatestTotalVolumeTicker As String
    Dim WorkSheetNames As Variant
    WorkSheetNames = Array("A", "B", "C", "D", "E", "F")
    
    For Each wsName In WorkSheetNames
        
        counter = 2
        
        Set Worksheet = ThisWorkbook.Worksheets(wsName)
        With Worksheet
        
            'Last Row'
            LR = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            GreatestPercentIncrease = 0
            GreatestPercentDecrease = 0
            GreatestTotalVolume = 0
            
            For i = 2 To LR
                If .Cells(i, 1).Value <> .Cells(i - 1, 1).Value Then
                    OpenPrice = .Cells(i, 3).Value
                    TotalVolume = 0
                End If
                
                TotalVolume = TotalVolume + .Cells(i, 7).Value
                
                If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Or i = LR Then
                    Ticker = .Cells(i, 1).Value
                    ClosePrice = .Cells(i, 6).Value
                    
                    YearlyChange = ClosePrice - OpenPrice
                    If OpenPrice <> 0 Then
                        PercentChange = (YearlyChange / OpenPrice) * 100
                    Else
                        PercentChange = 0
                    End If
                    
                    
                    
                    .Cells(counter, 9).Value = Ticker
                    .Cells(counter, 10).Value = YearlyChange
                    .Cells(counter, 11).Value = PercentChange & "%"
                    .Cells(counter, 12).Value = TotalVolume
                    
                    If PercentChange > GreatestPercentIncrease Then
                        GreatestPercentIncrease = PercentChange
                        GreatestPercentIncreaseTicker = Ticker
                    End If
                    If PercentChange < GreatestPercentDecrease Then
                        GreatestPercentDecrease = PercentChange
                        GreatestPercentDecreaseTicker = Ticker
                    End If
                    
                    If TotalVolume > GreatestTotalVolume Then
                        GreatestTotalVolume = TotalVolume
                        GreatestTotalVolumeTicker = Ticker
                    End If
                    
                    counter = counter + 1
                End If
            Next i
            
            .Cells(2, 14).Value = "Greatest % Increase"
            .Cells(3, 14).Value = "Greatest % Decrease"
            .Cells(4, 14).Value = "Greatest Total Volume"
            .Cells(1, 15).Value = "Ticker"
            .Cells(1, 16).Value = "Value"
            
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Volume"
            
            .Cells(2, 15).Value = GreatestPercentIncreaseTicker
            .Cells(2, 16).Value = GreatestPercentIncrease & "%"
            .Cells(3, 15).Value = GreatestPercentDecreaseTicker
            .Cells(3, 16).Value = GreatestPercentDecrease & "%"
            .Cells(4, 15).Value = GreatestTotalVolumeTicker
            .Cells(4, 16).Value = GreatestTotalVolume
        End With
    Next wsName
End Sub


