Sub stock_summary()

  ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
  ' --------------------------------------------
 
 Dim ws As Worksheet

    
 For Each ws In ThisWorkbook.Worksheets
 
   ' Set an initial variable for holding the ticker name
  Dim tickerName As String

  ' Set an initial variable for holding the total per ticker name
  Dim tickerTotal As Double
  tickerTotal = 0

  ' Keep track of the location for each ticker in the summary table
  Dim summaryTableRow As Integer
  summaryTableRow = 2
  
  'Create variables for yearly change and percent change
  Dim yearChange As Double
  Dim percentChange As Double
  
  'Create variables for open, closing price to be used in yearly change
  Dim openStockPrice As Double
  Dim closeStockPrice As Double
    
  'starting point for open stock price
  openStockPrice = Cells(2, 3).Value
  'Dim max As Double
          
  ws.Activate
 
  ' Determine the ticker Last Row
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stock tickers
  For I = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
      'Set close price
      closeStockPrice = Cells(I, 6).Value
      
      ' Set the ticker name
      tickerName = ws.Cells(I, 1).Value
            
      ' Print the ticker Name in the Summary Table
      Range("J" & summaryTableRow).Value = tickerName
      Range("J1").Value = "Ticker"
      
      'Calculate the yearly change in stock price
      yearChange = closeStockPrice - openStockPrice
      
      ' Print the yearly change in the Summary Table
      Range("K" & summaryTableRow).Value = yearChange
      Range("K" & summaryTableRow).NumberFormat = "0.00"
      Range("K1").Value = "Yearly Change"
      
      ' Calculate Percent Change accounting for the denominator being zero
      Range("L1").Value = "Percent Change"
      
      If (openStockPrice = 0 And closeStockPrice = 0) Then
         percentChange = 0
      
      ElseIf (openStockPrice = 0 And closeStockPrice <> 0) Then
         percentChange = 1
      
      Else
         percentChange = yearChange / openStockPrice
         Range("L" & summaryTableRow).Value = percentChange
         Range("L" & summaryTableRow).NumberFormat = "0.00%"
      
      End If
      
      ' Add to the ticker Total
      tickerTotal = tickerTotal + ws.Cells(I, 7).Value
      
      ' Print the ticker volume to the Summary Table
      Range("M" & summaryTableRow).Value = tickerTotal
      Range("M1").Value = "Total Stock Volume"

      ' Add one to the summary table row
      summaryTableRow = summaryTableRow + 1
      
      ' Reset the ticker Total and open stock price
      tickerTotal = 0
      openStockPrice = Cells(I + 1, 3)
      
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the ticker Total
      tickerTotal = tickerTotal + Cells(I, 7).Value

    End If

  Next I
  
  'Color the yearly change based on positive or negative value by:
  
  'Determine the yearly change last row
  yearChangeLR = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
  'Color cells by iterating through column 11 for positive or negative values
  For j = 2 To yearChangeLR
       
        If (Cells(j, 11).Value > 0 Or Cells(j, 11).Value = 0) Then
            
            Cells(j, 11).Interior.ColorIndex = 10
        
        ElseIf Cells(j, 11).Value < 0 Then
              
            Cells(j, Column + 11).Interior.ColorIndex = 3
        
        End If
        
        Next j
        
  'Create titles for Greatest % increase, % decrease, total volume, ticker and value
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"
            
  For k = 2 To yearChangeLR
        If ws.Cells(k, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & yearChangeLR)) Then
            
            Max = ws.Cells(k, 12).Value
            Range("Q2").Value = Max
            Range("Q2").NumberFormat = "0.00%"
            tickerName = ws.Cells(k, 10).Value
            Range("P2").Value = tickerName
            
        ElseIf ws.Cells(k, 12).Value = WorksheetFunction.Min(ws.Range("L2:L" & yearChangeLR)) Then
            
            Min = ws.Cells(k, 12).Value
            Range("Q3").Value = Min
            Range("Q3").NumberFormat = "0.00%"
            tickerName = ws.Cells(k, 10).Value
            Range("P3").Value = tickerName
            
        ElseIf ws.Cells(k, 13).Value = WorksheetFunction.Max(ws.Range("M2:M" & yearChangeLR)) Then
            
            Max = ws.Cells(k, 13).Value
            Range("Q4").Value = Max
            Range("Q4").NumberFormat = "0"
            tickerName = ws.Cells(k, 10).Value
            Range("P4").Value = tickerName
        
        End If
        
  Next k
        
 'Auto adjust the column widths in each worksheet so the summary table data is readable
 ws.Columns("J:Q").AutoFit
        
Next ws

 ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    'move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")
    
    Const excludeSheets As String = "StockCalculations"
        
    ' Loop through all sheets
        Dim sh As Worksheet
        
        For Each sh In ThisWorkbook.Worksheets
        
        sh.Activate
        
         If (sh.Name = (excludeSheets)) Then
        
        sh.Range("J:Q").ClearContents
        sh.Range("J:K").Interior.Color = xlNone
        
        Else
                        
        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastrow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        lastRowYr = sh.Cells(Rows.Count, "A").End(xlUp).Row - 1
                
        ' Copy the contents of each sheet into the combined sheet
        combined_sheet.Range("A" & lastrow & ":D" & ((lastRowYr - 1) + lastrow)).Value = sh.Range("J2:M" & (lastRowYr + 1)).Value
        
        End If
                
        If (sh.Name = (excludeSheets)) Then
        
        sh.Range("J:Q").ClearContents
        sh.Range("J:K").Interior.Color = xlNone
        
        Else
        ' Copy the headers from sheet 1
        combined_sheet.Range("A1:D1").Value = Sheets(2).Range("J1:M1").Value
                                
        ' Autofit to display data
        combined_sheet.Columns("A:D").AutoFit
        
        End If
                                
        Next sh

End Sub

Sub clearStockSummary()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

 ws.Activate
 
 ws.Range("J:Q").ClearContents
 ws.Range("J:K").Interior.Color = xlNone
 
If ws.Name = "Combined_Data" Then

Application.DisplayAlerts = False

Sheets("Combined_Data").Delete

Application.DisplayAlerts = True

End If
 
Next ws

End Sub

