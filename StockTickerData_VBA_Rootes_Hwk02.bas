Attribute VB_Name = "Module2"
Sub LoopThroughTabs()

    ' Declare a variable for the current tab (worksheet)
    Dim currentTab As Worksheet

    ' Loop through all tabs (worksheets)
    For Each currentTab In Worksheets

        currentTab.Select
        LoopThroughDailyStockData (currentTab.name)
         
    Next

End Sub

Sub LoopThroughDailyStockData(name As String)

    ' Declare Stock Ticker Variables
    ' <ticker>    <date>  <open>  <high>  <low>   <close> <vol>
    Dim ticker As String
    Dim tickerDate As Date
    Dim openPrice, highPrice, lowPrice, closePrice As Double
    Dim volume, row, resultsRow As Double
    Dim tickerCol, dataStartRow, dateCol, openCol, highCol, lowCol, closeCol, volCol As Integer
    
    
    ' Declare calculation and by ticker tracking variables
    Dim startPrice, endPrice, yearlyChange, yearlyPercentChange As Double
    Dim totalVolume As Double
    Dim startDate, endDate As Date
    Dim headerRow, totalVolCol, yearlyChangeCol, percentChangeCol As Integer
    Dim tickerSummaryCol As Integer
        
    
    ' Declare overall tracking variables
    Dim greatestPercentIncrease, greatestPercentDecrease, greatestTotalVolume As Double
    Dim greatestPercentIncrTicker, greatestPercentDecrTicker, greatestTotalVolTicker As String
        
    
    ' Initialize variables
    headerRow = 1
    tickerCol = 1
    dateCol = 2
    openCol = 3
    highCol = 4
    lowCol = 5
    closeCol = 6
    volCol = 7
    dataStartRow = 2
    tickerSummaryCol = 9
    yearlyChangeCol = 10
    percentChangeCol = 11
    totalVolCol = 12
    ticker = ""
    
    Cells(headerRow, tickerSummaryCol).Value = "Ticker"
    Cells(headerRow, yearlyChangeCol).Value = "Yearly Change"
    Cells(headerRow, percentChangeCol).Value = "Percent Change"
    Cells(headerRow, totalVolCol).Value = "Total Volume"
    Cells(headerRow, 16).Value = "Ticker"
    Cells(headerRow, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    resultsRow = 2
    row = dataStartRow
    
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    
    
    ' Loop through rows:  Note row 1 has headers
    ' NOTE:  This solution assumes all ticker data is grouped and in order from oldest to newest
     Do While Not IsEmpty(Cells(row, tickerCol))
    
        ' If the ticker symbol is the same as the last row, update values
        If (ticker = Cells(row, tickerCol)) Then
        
            endPrice = Cells(row, closeCol).Value
            totalVolume = totalVolume + Cells(row, volCol).Value

        End If
        
        
        ' If this is not the starting row of data and then next row ticker does not match, then process current ticker results
        If (row <> 2 And Cells(row + 1, tickerCol).Value <> Cells(row, tickerCol).Value) Then
        
            ' Calculate Ticker Yearly Metrics
            yearlyChange = endPrice - startPrice
            If (startPrice = 0) Then
                yearlyPercentChange = 0
            Else
                yearlyPercentChange = yearlyChange / startPrice
            End If
            
            ' Update Spreadsheet with Ticker Metrics
            Cells(resultsRow, tickerSummaryCol).Value = ticker
            Cells(resultsRow, yearlyChangeCol).Value = yearlyChange
            Cells(resultsRow, percentChangeCol).Value = yearlyPercentChange
            Cells(resultsRow, totalVolCol).Value = totalVolume
            
            ' Color code yearly changes
            If (yearlyChange > 0) Then
                Cells(resultsRow, yearlyChangeCol).Interior.Color = vbGreen
            End If
            
            If (yearlyChange < 0) Then
                Cells(resultsRow, yearlyChangeCol).Interior.Color = vbRed
            End If
                            
            ' increment row for the next ticker results
            resultsRow = resultsRow + 1
            
            ' determine if the current ticker metrics exceed current tracking metrics
            If (greatestPercentIncrease < yearlyPercentChange) Then
                greatestPercentIncrease = yearlyPercentChange
                greatestPercentIncrTicker = ticker
            End If
            
            If (greatestPercentDecrease > yearlyPercentChange) Then
                greatestPercentDecrease = yearlyPercentChange
                greatestPercentDecrTicker = ticker
            End If
            
            If (greatestTotalVolume < totalVolume) Then
                greatestTotalVolume = totalVolume
                greatestTotalVolTicker = ticker
            End If
        
        End If
                
        ' If the current ticker does not match the current row, update the ticker and set initial values
        If (ticker <> Cells(row, tickerCol).Value) Then
            ticker = Cells(row, tickerCol).Value
            startPrice = Cells(row, openCol).Value
            totalVolume = Cells(row, volCol).Value
        End If
        
        ' Increment the row to read the next result
        row = row + 1
    Loop
    
    ' Update overall metrics for the page
    Cells(2, 16).Value = greatestPercentIncrTicker
    Cells(2, 17).Value = greatestPercentIncrease
    Cells(3, 16).Value = greatestPercentDecrTicker
    Cells(3, 17).Value = greatestPercentDecrease
    Cells(4, 16).Value = greatestTotalVolTicker
    Cells(4, 17).Value = greatestTotalVolume
    
    ' Format the Data
    FormatTickerTab (name)
    
End Sub

Sub FormatTickerTab(name As String)

    SetColumnWidths

    AlignDataWithinCells

    SetDataType
    
    AddGridLines
    
    FormatHeaders
    
    AddYearToMetricsTable (name)

End Sub

Sub SetColumnWidths()

    ' Set column widths
    Columns("A:G").Select
    Selection.ColumnWidth = 10
    
    Columns("I:L").Select
    Selection.ColumnWidth = 15
    
    Columns("P:P").Select
    Selection.ColumnWidth = 10
     
    Columns("Q:Q").Select
    Selection.ColumnWidth = 19
    
    Columns("O:O").Select
    Selection.ColumnWidth = 19

End Sub

Sub AlignDataWithinCells()

    ' Align data within cells
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
End Sub
Sub SetDataType()

    ' Format data as currency
    Range("C2:F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    
    Range("J2:J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    
    ' Format data as percent
    Range("K2:K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00%"
    
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"

End Sub

Sub AddGridLines()
    
    ' Ticker Data
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    SetGridLines
    
    ' Resuls by Ticker Data
    Range("I1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    SetGridLines
    
    ' Overall Metrics
    Range("O1:Q4").Select
    SetGridLines
    
End Sub
Sub SetGridLines()
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

Sub FormatHeaders()
    
    Range("P1:Q1,O2:O4,I1:L1,A1:G1").Select
    Range("G1").Activate
    FormatTitles
    
End Sub

Sub FormatTitles()

    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
End Sub

Sub AddYearToMetricsTable(name As String)

    Cells(1, 15).Value = name
    Range("O1:O1").Select
    Selection.Font.Bold = True
    With Selection.Font
        .name = "Calibri"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

End Sub

