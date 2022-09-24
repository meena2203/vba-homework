Attribute VB_Name = "Module1"

Public Sub StockMarketAnalyst()
    'Step 1: Key assumption for the loop - Data is sorted by ticker and date
    
    'Step 2: Declare and define variables
    Dim InputTicker, OutputTicker As String
    Dim OpenPrice, High, Low, ClosePrice, YearlyChange, PercentChange, GreatestPercentIncrease, GreatestPercentDecrease, GreatestTotalVolume As Double
    Dim StartRowNumber, LastRowNumber, StockVolume, TotalStockVolume, OutputRowNumber As Integer
    Dim YearlyChangeColumn, PercentChangeColumn As Range
    
    StartRowNumber = 2
    LastRowNumber = Range("A1").End(xlDown).Row
    OutputRowNumber = 1
    
    InputTicker = Cells(StartRowNumber, 1)
    OpenPrice = Cells(StartRowNumber, 3)
    ClosePrice = Cells(StartRowNumber, 6)
    StockVolume = Cells(StartRowNumber, 7)
    
    'Step 3: Print header for the output data
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Step 4: Start the loop for the output data (for last row data to be analysed loop must go to the next empty row, activate the "new sticker" code and finish calculations for the last row)
    For i = StartRowNumber + 1 To LastRowNumber + 1
        
        'Step 4a: if same ticker
        If Cells(i, 1).Value = Cells(i - 1, 1) Then
                StockVolume = StockVolume + Cells(i, 7).Value
          
            'Step 4b: if Not same ticker
            Else
                ClosePrice = Cells(i - 1, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                PercentChange = YearlyChange / OpenPrice
                
                'Step 4c: Print in output table
                OutputRowNumber = OutputRowNumber + 1
                Cells(OutputRowNumber, 9).Value = InputTicker
                Cells(OutputRowNumber, 10).Value = YearlyChange
                Cells(OutputRowNumber, 11).Value = PercentChange
                Cells(OutputRowNumber, 12).Value = StockVolume
                
                'Step 4d: Reset values for the new inputticker
                InputTicker = Cells(i, 1)
                OpenPrice = Cells(i, 3)
                StockVolume = Cells(i, 7)
        End If
        
    Next i
    
    
    'Step 5. Conditional formatting for yearlychangecolumn
    Set YearlyChangeColumn = Range(Cells(2, 10), Cells(OutputRowNumber, 10))
    
    YearlyChangeColumn.FormatConditions.Delete
    
    YearlyChangeColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
    YearlyChangeColumn.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
    
    YearlyChangeColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
    YearlyChangeColumn.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
    
    'Step 6. Conditional formatting for percentchangecolumn
    Set PercentChangeColumn = Range(Cells(2, 11), Cells(OutputRowNumber, 11))
    
    PercentChangeColumn.FormatConditions.Delete
    
    PercentChangeColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
    PercentChangeColumn.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
    
    PercentChangeColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
    PercentChangeColumn.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
    
    
    'Step 7.Format output columns
    Columns(9).Select
    Columns(9).EntireColumn.AutoFit
    Columns(10).Select
    Columns(10).EntireColumn.AutoFit
    Selection.Style = "Comma"
    Selection.NumberFormat = "0.00"
    Columns(11).Select
    Columns(11).EntireColumn.AutoFit
    Selection.NumberFormat = "0.00%"
    Columns(12).Select
    Columns(12).EntireColumn.AutoFit
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Cells(2, 10).Select
    
    
    'Bonus Question
    'Step 8. Print header for bonus table
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Step 9. Print Parameters for bonus table
    Cells(2, 15).Value = "Greatest % Increse"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    'Step 10. Print Values for bonus table
    GreatestPercentIncrease = Application.WorksheetFunction.Max(Columns(11))
    GreatestPercentDecrease = Application.WorksheetFunction.Min(Columns(11))
    GreatestTotalVolume = Application.WorksheetFunction.Max(Columns(12))
    
    Cells(2, 17).Value = GreatestPercentIncrease
    Cells(3, 17).Value = GreatestPercentDecrease
    Cells(4, 17).Value = GreatestTotalVolume
    
    'Step 11.Loop to find the ticker for the bonus item
    For i = 2 To OutputRowNumber
        
        If Cells(i, 11).Value = GreatestPercentIncrease Then
            greatestPercentIncreaseTicker = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value = GreatestPercentDecrease Then
            greatestPercentDecreaseTicker = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value = GreatestTotalVolume Then
            greatesttotalvolumeTicker = Cells(i, 9).Value
        End If
    
    Next i
    
    Cells(2, 16).Value = greatestPercentIncreaseTicker
    Cells(3, 16).Value = greatestPercentDecreaseTicker
    Cells(4, 16).Value = greatesttotalvolumeTicker
    
    'Step 12. Format bonus table
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns(15).Select
    Columns(15).EntireColumn.AutoFit
    Columns(16).Select
    Columns(16).EntireColumn.AutoFit
    Columns(17).Select
    Columns(17).EntireColumn.AutoFit
    Cells(2, 9).Select
    
End Sub
