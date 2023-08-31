Attribute VB_Name = "Module2"
Sub SummaryForMultipleSheets()
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
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
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
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""
        
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
        
        ' Loop through unique ticker symbols to calculate metrics
        For Each currentTicker In uniqueTickers
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
            
            ' Calculate yearly change and percentage change
            If openingPrice <> 0 Then
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentageChange = (yearlyChange / openingPrice) * 100
                End If
            End If
            
            ' Update greatest metrics and corresponding tickers
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                greatestIncreaseTicker = currentTicker
            ElseIf percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                greatestDecreaseTicker = currentTicker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = currentTicker
            End If
        Next currentTicker
        
        ' Output greatest metrics and corresponding tickers
        ws.Cells(2, outputColumn + 7).Value = greatestIncreaseTicker
        ws.Cells(2, outputColumn + 8).Value = greatestIncrease
        ws.Cells(3, outputColumn + 7).Value = greatestDecreaseTicker
        ws.Cells(3, outputColumn + 8).Value = greatestDecrease
        ws.Cells(4, outputColumn + 7).Value = greatestVolumeTicker
        ws.Cells(4, outputColumn + 8).Value = greatestVolume
    Next sheetName
End Sub
