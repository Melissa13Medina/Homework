Sub TickerTapeParade()

    Dim i As Long

    Dim LastRow As Long
    Dim UniqueRow As Integer
    Dim CurrentStock As String
    Dim CurrentVolume As Integer
    Dim NewStock As String
    Dim TotalRow As Integer

    LastRow = cells(Rows.Count, 1).End(xlUp).Row

    Dim TotalStockVolume As Integer
    Dim TotalStockName As String
    
    TotalRow = 0

    For i = 2 To LastRow  
        
        CurrentStock = Cells(i, 1).Value
        CurrentVolume = Cells(i, 7).Value
        NewStock = Cells(i + 1, 1).Value
              
        TotalRow = TotalRow + CurrentVolume
        
        If CurrentStock <> NewStock Then
            UniqueRow = UniqueRow + 1
            
        Range("I" & UniqueRow + 1).Value = CurrentStock
        Range("L" & UniqueRow + 1).Value = TotalRow
                   
           TotalRow = 0
        End If
        
    Next i
    Range("I1").value = "Ticker"
    Range("J1").value = "Total Stock Volume"
    Range("I:J").EntireColumn.AutoFit

End Sub