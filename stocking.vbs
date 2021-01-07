Sub Stocking()
    Dim countTicks As Integer
    Dim opening As Double
    Dim total As Double
    Dim greatIncreaseName As String
    Dim greatIncreaseValue As Double
    Dim greatDecreaseName As String
    Dim greatDecreaseValue As Double
    Dim greatTotalVolumeName As String
    Dim greatTotalVolumeValue As Double
    
    Dim WS_Count As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For K = 1 To WS_Count
        ActiveWorkbook.Worksheets(K).Activate
        total = 0
        opening = Cells(2, 3).Value
        countTicks = 2
        
        greatIncreaseName = ""
        greatIncreaseValue = 0
        greatDecreaseName = ""
        greatDecreaseValue = 0
        greatTotalVolumeName = ""
        greatTotalVolumeValue = 0
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % increase"
        Cells(3, 15).Value = "Greatest % decrease"
        Cells(4, 15).Value = "Greates Total Volume"
        
        
        For i = 2 To Range("A1", Range("A1").End(xlDown)).Count
            total = total + Cells(i, 7)
            If (Cells(i, 1) <> Cells(i + 1, 1)) Then
                Cells(countTicks, 9).Value = Cells(i, 1).Value
                Cells(countTicks, 10).Value = Cells(i, 6).Value - opening
                If (opening <> 0) Then
                    Cells(countTicks, 11).Value = (Cells(countTicks, 10).Value) / opening
                Else
                    Cells(countTicks, 11).Value = 0
                End If
                Cells(countTicks, 11).NumberFormat = "0.00%"
                Cells(countTicks, 12).Value = total
                If (Cells(countTicks, 10).Value > 0) Then
                    Cells(countTicks, 10).Interior.ColorIndex = 4
                Else
                    Cells(countTicks, 10).Interior.ColorIndex = 3
                End If
                If (Cells(countTicks, 11).Value > greatIncreaseValue) Then
                    greatIncreaseValue = Cells(countTicks, 11).Value
                    greatIncreaseName = Cells(countTicks, 9).Value
                End If
                If (Cells(countTicks, 11).Value < greatDecreaseValue) Then
                    greatDecreaseValue = Cells(countTicks, 11).Value
                    greatDecreaseName = Cells(countTicks, 9).Value
                End If
                If (total > greatTotalVolumeValue) Then
                    greatTotalVolumeValue = total
                    greatTotalVolumeName = Cells(countTicks, 9).Value
                End If
                
                opening = Cells(i + 1, 3).Value
                countTicks = countTicks + 1
                total = 0
            End If
        Next i
        
        Cells(2, 16).Value = greatIncreaseName
        Cells(2, 17).Value = greatIncreaseValue
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = greatDecreaseName
        Cells(3, 17).Value = greatDecreaseValue
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 16).Value = greatTotalVolumeName
        Cells(4, 17).Value = greatTotalVolumeValue
    Next K
End Sub


