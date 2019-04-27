Attribute VB_Name = "Module1"


'Sub atttached to my Calculate button
Sub RunCode()

'declare my variables
Dim max_tic As String
Dim min_tic As String
Dim max_vol_tic As String
Dim TickerSymbol As String
Dim current_sheet As String
Dim i As Long
Dim j As Long
Dim OutputRow As Integer
Dim previous_cell As Integer
Dim YR_Closing_Price As Double
Dim YR_Opening_Price As Double
Dim YR_Change As Double
Dim YR_Percent_Change As Double
Dim TotalVolume As Double
Dim max As Double
Dim min As Double
Dim max_match As Double
Dim min_match As Double
Dim Max_T_Vol As Double
Dim GT_Percent_Increase As String
Dim GT_Percent_Decrease As String
Dim New_Array_Incre() As String
Dim New_Array_Decre() As String
Dim Last_Row As Long
Dim LR As Long
Set Sheet = ThisWorkbook.Sheets





'set the last row
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row


'set counter for total volume
TotalVolume = 0


'ouput table row counter set.  I believe this is a redundancy since I set this once I activated my sheets in the parent for loop below,_
'but I just left it in.  If it ain't broke, don't try to fix it and break the rest....
OutputRow = 2

'set ticker symbol to null
TickerSymbol = ""








'parent For Each loop to iterate through sheets
For Each Sheet In Sheets
  
  
  
  'activate current seet
  Sheet.Activate
  
  'reset row counter to 2 for my output table so I start under the titles of the output table as I traverse through sheets
     OutputRow = 2
     
  'set the titles for the output tables
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
  
  'format YR Change column and percentage change columns accordingly for the output table
  Sheet.Range("J:J").NumberFormat = "0.00000000"
  Sheet.Range("K:K").NumberFormat = "0.00%"
  
  
  'child For Loop to itterate through rows in each sheet
  For i = 2 To Last_Row
  
    'If Condition to check for opening price
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
      
      'if the condition checks out... set opening price to the current cell
      YR_Opening_Price = Cells(i, 3).Value
    
    
    
    'If Condition to check if the next cell matched the current one.  Also have to have a condition to make sure I dont try to divide by zero for_
    'the percent change calculations, as without this, I will throw an overflow error
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And YR_Opening_Price > 0 Then
     
      'add current Ticker's volume to the total
      TotalVolume = TotalVolume + Cells(i, 7).Value
     
      'Define the stock ticker
      TickerSymbol = Cells(i, 1).Value
     
      'set yearly closing price
      YR_Closing_Price = Cells(i, 6).Value
      
      'do calculations for yearly change and yearly percent change
      YR_Change = (YR_Opening_Price - YR_Closing_Price) * -1
      YR_Percent_Change = (YR_Closing_Price - YR_Opening_Price) / YR_Opening_Price
     
  
     
     
      'populate Ticker column in table
      Cells(OutputRow, 9).Value = TickerSymbol
     
     
      'populate Yearly Change column in output table
      Cells(OutputRow, 10).Value = YR_Change
      
     
      'populate Precent Change column in output table
      Cells(OutputRow, 11).Value = YR_Percent_Change
     
      'poplulate Total Volume column in output table
      Cells(OutputRow, 12).Value = TotalVolume
      
      'increase counter by one so I can print out the info for the next unique ticker in my output table
      OutputRow = OutputRow + 1
     
     
      'reset total volume counter
      TotalVolume = 0
     
     
     
     
      Else
     
        'add current volume to toatal if next cell for ticker symbol matches
        TotalVolume = TotalVolume + Cells(i, 7).Value
     
     
    End If
     

  Next i
  
  
  
  
  
  
  
  
  
    'new last row declaration so we just loop to the last row of our output table range and increase the efficiency/ cleaness of our program
    LR = Sheet.Cells(Sheet.Rows.Count, "J").End(xlUp).Row
    
    
    
    'Make titles for our secondary output table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    
   'format secondary output table properly for GT % Increase and GT % Decrease
    Sheet.Range("Q2", "Q3").NumberFormat = "0.00%"
    
    
    'created another For Loop outside of my previous loop so I could format the colors for yearly change in between the iteration to the next sheet
    For j = 2 To LR
    
    
       'If stock ended higher or had no channge for the year, color cell green.  otherwise, it's a red cell...
       If Cells(j, 10) >= 0 Then
      
       Cells(j, 10).Interior.ColorIndex = 4
      
       ElseIf Cells(j, 10) < 0 Then
      
       Cells(j, 10).Interior.ColorIndex = 3
       
      
       End If
       
     
    Next j
    
    
   

    'find my min and max value for the percent change and my max total volume change
    max = Application.WorksheetFunction.max(Range("k:k"))
    min = Application.WorksheetFunction.min(Range("k:k"))
    Max_T_Vol = Application.WorksheetFunction.max(Range("l:l"))
    
    
    'matching algo to set a variable to find the corresponding ticker symbol for min and max
    max_match = Application.WorksheetFunction.Match(Application.WorksheetFunction.max(Range("k:k")), Range("k:k"), 0)
    min_match = Application.WorksheetFunction.Match(Application.WorksheetFunction.min(Range("k:k")), Range("k:k"), 0)
    vol_match = Application.WorksheetFunction.Match(Application.WorksheetFunction.max(Range("l:l")), Range("l:l"), 0)
    
    
    'find the ticker symbol for my min and max in the first column
    max_tic = Range("i:i")(max_match)
    min_tic = Range("i:i")(min_match)
    max_vol_tic = Range("i:i")(vol_match)
    
       
    
    'populate my secondary output table with the values
    Cells(2, 16).Value = max_tic
    Cells(2, 17).Value = max
    
    Cells(3, 16).Value = min_tic
    Cells(3, 17).Value = min
    
    Cells(4, 16).Value = max_vol_tic
    Cells(4, 17).Value = Max_T_Vol
    
  
  
Next Sheet


End Sub





