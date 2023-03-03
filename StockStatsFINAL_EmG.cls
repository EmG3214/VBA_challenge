VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockStatsFINAL()
'loop through each worksheet
For Each ws In Worksheets

  ' Set an initial variable for holding the ticker symbol and one for holding the amount of tickers in a set
  Dim ticker As String
Dim tickersum As Double
tickersum = 0

  ' Set an initial variable for holding the total volume sum per ticker symbol
  Dim VolumeSum As Double
  VolumeSum = 0
  
  
  'Set an initial variable for holding the close number for later average and difference calculations
  Dim CloseNum As Double
  CloseNum = 0
  
  'Set an initial variable for holding the open number for later average and difference calculations
  Dim OpenNum As Double
  OpenNum = 0
 
' Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

 'do the lastrow formula to find the lasr row on each sheet
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


  ' Loop through all stock data
  Dim i As Double
  
  For i = 2 To LastRow

    ' Check if we are still within the same ticker symbol, if it is NOT equal...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker symbol
      ticker = ws.Cells(i, 1).Value

      ' Add to the Volume summary
      VolumeSum = VolumeSum + ws.Cells(i, 7).Value
      
      'where the tickers are different, closeNum is the last close number in that list and is saved here
      CloseNum = ws.Cells(i, 6).Value
      
      'this counts the number of tickers in a set
      tickersum = tickersum + ws.Cells(i, 1).Count
      
      'this takes the tickersum and finds the first opening value of the set
      OpenNum = ws.Cells((i - tickersum + 1), 3).Value
      
      ' Print the Ticker Symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Print the Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = VolumeSum

      'Calculate and Print the year change number
      ws.Range("J" & Summary_Table_Row).Value = CloseNum - OpenNum
      
      'calculate the percentage change number and print the percentage
      ws.Range("K" & Summary_Table_Row).Value = ((CloseNum - OpenNum) / OpenNum)
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume sum for the next loop
        VolumeSum = 0
        
        'reset the ticer count for the new set
        tickersum = 0

    ' If the cell immediately following a row is the same...
    Else

      ' Add to the Volume Sum
      VolumeSum = VolumeSum + ws.Cells(i, 7).Value
      
      'Add to the ticker counter for finding the beginning year open price
      tickersum = tickersum + ws.Cells(i, 1).Count

    End If
Next i

'Create Headers for Summary Table
ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Price Change"
ws.Range("K1").Value = "Yearly Percent Change"
ws.Range("L1").Value = "Sum of Volume"


'----------------------------------------------

'set summary table last row
newLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Add in conditional formatting for summary table for both yearly change number and percentage

For i = 2 To newLastRow
    
  For j = 10 To 11
        If ws.Cells(i, j).Value > 0 Then
            ws.Cells(i, j).Interior.ColorIndex = 4
         Else
            ws.Cells(i, j).Interior.ColorIndex = 3
         End If
    Next j

Next i

'hard challenge addding in the greatest increase, decrease, and total volume to each sheet

Dim max As Double
Dim min As Double
Dim maxvol As Double
Dim tag1 As String
Dim tag2 As String
Dim tag3 As String

tag1 = ""
tag2 = ""
tag3 = ""
max = 0
min = 0
maxvol = 0

For i = 2 To newLastRow
'if the item is greater than the set max, then save it and the ticker symbol to that variables
    If ws.Cells(i, 11) >= max Then
        max = ws.Cells(i, 11)
        tag1 = ws.Cells(i, 9)
    End If
    
 'if the item is less than the set minimum, then set the new value as the min and save the ticker symbol to tag
    If ws.Cells(i, 11) <= min Then
        min = ws.Cells(i, 11)
        tag2 = ws.Cells(i, 9)
    End If
    
    'if the item is greater than the set max, then save it and the ticker symbol to that variables
    If ws.Cells(i, 12) >= maxvol Then
        maxvol = ws.Cells(i, 12)
        tag3 = ws.Cells(i, 9)
    End If
    
Next i

'add labels and paste in data
ws.Cells(1, 16).Value = "Value"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

ws.Cells(2, 16).Value = max
ws.Cells(2, 15).Value = tag1

ws.Cells(3, 16).Value = min
ws.Cells(3, 15).Value = tag2

ws.Cells(4, 16).Value = maxvol
ws.Cells(4, 15).Value = tag3



'add in formatting for showing columk K as percentages, and other values with appropriate decimal places and values

'set percent change column as a percentage with appropriate decimal places
ws.Range("J:J").NumberFormat = "0.00"
ws.Range("K:K").NumberFormat = "0.00%"
ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(3, 16).NumberFormat = "0.00%"
ws.Cells(4, 16).NumberFormat = "0,000.00"

'end loop the code through all sheets and send a message that we are done

Next ws

    MsgBox ("Completed all sheets")

End Sub

