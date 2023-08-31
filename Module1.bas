Attribute VB_Name = "Module1"
Sub LoopThroughStocksForMultipleSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Long
    Dim openingPriceColumn As Long
    Dim closingPriceColumn As Long
    Dim volumeColumn As Long
    Dim outputColumn As Long
    Dim dateColumn As Long
    Dim currentTicker As Variant
    Dim outputRow As Long
    Dim uniqueTickers As Collection
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim sheetName As Variant
    
    ' Define an array of sheet names to loop through
    Dim sheetNames As Variant
    sheetNames = Array("2018", "2019", "2020")
    
    For Each sheetName In sheetNames
        ' Set the worksheet to work with
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Define the column numbers for relevant data
        tickerColumn = 1 ' Assuming ticker symbols are in column A
        dateColumn = 2    ' Date column
        openingPriceColumn = 3 ' Opening price column
        closingPriceColumn = 6 ' Closing price column
        volumeColumn = 7 ' Volume column
        outputColumn = 9 ' Column where the output will be placed
        
        ' Initialize variables
        Set uniqueTickers = New Collection
        outputRow = 2 ' Start output from row 2
        
        ' Find the last row with data
        lastRow = ws.Cells(ws.Rows.Count, tickerColumn).End(xlUp).Row
        
        ' Loop through each row of data
        For currentRow = 2 To lastRow
            currentTicker = ws.Cells(currentRow, tickerColumn).Value
            ' Check if the ticker symbol is unique
            On Error Resume Next
            uniqueTickers.Add currentTicker, CStr(currentTicker)
            On Error GoTo 0
        Next currentRow
        
        ' Output unique ticker symbols, calculate yearly and percentage change
        For Each currentTicker In uniqueTickers
            ws.Cells(outputRow, outputColumn).Value = currentTicker
            openingPrice = 0 ' Reset opening price for each ticker
            yearlyChange = 0
            percentageChange = 0
            totalVolume = 0
            
            ' Loop through rows to calculate yearly change, percentage change, and total volume
            For currentRow = 2 To lastRow
                If ws.Cells(currentRow, tickerColumn).Value = currentTicker Then
                    If openingPrice = 0 Then
                        openingPrice = ws.Cells(currentRow, openingPriceColumn).Value
                    End If
                    closingPrice = ws.Cells(currentRow, closingPriceColumn).Value
                    totalVolume = totalVolume + ws.Cells(currentRow, volumeColumn).Value
                End If
            Next currentRow
            
            ' Calculate and output yearly change and percentage change
            If openingPrice <> 0 Then
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentageChange = (yearlyChange / openingPrice) * 100
                End If
            End If
            ws.Cells(outputRow, outputColumn + 1).Value = yearlyChange
            ws.Cells(outputRow, outputColumn + 2).Value = percentageChange
            ws.Cells(outputRow, outputColumn + 3).Value = totalVolume ' Output total volume
            
            outputRow = outputRow + 1
        Next currentTicker
        
        ' Apply conditional formatting...
Set Rng = ws.Range(ws.Cells(2, outputColumn + 1), ws.Cells(outputRow - 1, outputColumn + 1))
Rng.FormatConditions.AddColorScale ColorScaleType:=3
Rng.FormatConditions(Rng.FormatConditions.Count).SetFirstPriority
Rng.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValuePercentile
Rng.FormatConditions(1).ColorScaleCriteria(1).Value = 0
Rng.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
Rng.FormatConditions(1).ColorScaleCriteria(2).Value = 50
Rng.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)
Rng.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValuePercentile
Rng.FormatConditions(1).ColorScaleCriteria(3).Value = 90

' Apply colors for positive and negative changes
Rng.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red for lower values
Rng.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green for positive values

' Apply the same conditional formatting to Column K
Set RngK = ws.Range(ws.Cells(2, outputColumn + 2), ws.Cells(outputRow - 1, outputColumn + 2))
RngK.FormatConditions.AddColorScale ColorScaleType:=3
RngK.FormatConditions(RngK.FormatConditions.Count).SetFirstPriority
RngK.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValuePercentile
RngK.FormatConditions(1).ColorScaleCriteria(1).Value = 0
RngK.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
RngK.FormatConditions(1).ColorScaleCriteria(2).Value = 50
RngK.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)
RngK.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValuePercentile
RngK.FormatConditions(1).ColorScaleCriteria(3).Value = 90

' Apply colors for positive and negative changes to Column K
RngK.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) ' Red for lower values
RngK.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0) ' Green for positive values

    Next sheetName
End Sub


