Attribute VB_Name = "Module1"
Sub wallstreetsummaries()

'TO DO: Create a script that will loop through all the stocks for one year for each run and take the following information.
'An assumption I'm making with the data is that that it is sorted exactly the way I need it to be, which is by ticker(in ascending order from A-Z) and the dates are sorted in order from the earliest date to latest date.  The dates are formated in such a way that the are are also in sequential order from smallest number to largest number.  Given this, assumption proceed with the script

'The first thing I'll do here is to declare some variables for what is needed here.

'summary row count will be used to capture the row value for the last item in a particular group of data
Dim summaryRowCount As Integer

'Rowline will be used for my while loop as I traverse the data
'Dim rowline As Integer

'Next I'll create some variables to store the data points I'm looking to capture

  'The ticker symbol.
Dim sumTickerName As String

  'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
Dim sumPriceDiff As Double

  'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
Dim sumPercDiff As Double

  'The total stock volume of the stock.
Dim sumTotalStockVolume As Double

'I'm also going to create some variables to hold some values that I want to do calculations with to get to my sumPriceDiff and SumPerDiff
'firstPrice will hold the very first price in the group
Dim firstPrice As Double

'last price will hold the last price in the group.firstPrice
Dim lastPrice As Double



'Next, I'm going to set the values of these variables to zero, just as a matter of course to make sure they start out cleared.

sumTickerName = 0
sumPriceDiff = 0
sumPercDiff = 0
sumTotalStockVolume = 0

'I'm setting SummaryRowCount to 2 because I want the script to place the values in the second row of my spreadsheet, then move to the next rows from there
summaryRowCount = 2

'I 'll also set the firstPrice to the open value in my sheet
firstPrice = Cells(2, 3).Value


'Creating some Columns to store my caluculations
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'I'm going to do a little formatting of my cells to make sure that the data is more readable
Range("I1:L1").Font.Bold = True
Range("I1:L1").Font.Italic = True
Columns("I:L").AutoFit

'Now, I'll get into the main ask for this homework. I'll tackle the total volume first, since that is simpler
'For all of the rows in the spreadsheet, I want the program to check each row and then sum the values of the vol column

For rowline = 2 To Range("A2").End(xlDown).row

'It will take the value in the sumTotalStockVolume variable, then add the next value in the group, then re-assign it to that variable.
    sumTotalStockVolume = sumTotalStockVolume + Cells(rowline, 7).Value
 
'The script will do this until it detects a change in the ticker name type.  For each row, I will loop through each value in the first column until they aren't equal, or there's a change.
   If Cells(rowline, 1) <> Cells(rowline + 1, 1) Then
    
    'capture the name of the ticker type an place it in my sumTickerName variable
     sumTickerName = Cells(rowline, 1).Value
    
    'next, I'll place the value in my SumTickerName variable into the appropriate place in column L
    Cells(summaryRowCount, 9) = sumTickerName
    
    'capture the running total of the cc type an place it in the column H for total
    Cells(summaryRowCount, 12) = sumTotalStockVolume
    
    'capture the value of the last closing price and place it in my last prices variable
    lastPrice = Cells(rowline, 6).Value
    
    'capture the price difference between the opening price at the beginning of the year and the closing price at the end of the year
     sumPriceDiff = lastPrice - firstPrice
    
    'Place the contents of the sumPriceDiff into the appropriate cell
     Cells(summaryRowCount, 10) = sumPriceDiff
     
         'caputure the percentage difference between the opening prices at the beginning of the year and the closing price at the end of the year
         If firstPrice <> 0 Then
         'If lastPrice = 0 And firstPrice = 0 Then
         
         'sumPercDiff = 0
         
         'Else
         sumPercDiff = (lastPrice - firstPrice) / firstPrice
         
         End If
         
     'Place the contents of the SumPercDiff into the appropriate cell
     Cells(summaryRowCount, 11) = sumPercDiff
    
    
    'Reset the start value to the open value of the next group
    firstPrice = Cells(rowline + 1, 3)
    
    'increment summaryRowCount by 1 so that it moves on the to the next row
    summaryRowCount = summaryRowCount + 1
    
    'capture the price difference between the opening price at the beginning of the year and the closing price at the end of the year
    sumPriceDiff = lastPrice - firstPrice
    
    'Place the contents of the sumPriceDiff into the appropriate cell
    Cells(summaryRowCount, 10) = sumPriceDiff
    
    'test place the first value and the last value in the columns
    'Cells(summaryRowCount-1, 10) = firstPrice
    'Cells(summaryRowCount, 11) = lastPrice
    
    'clear the variables back to 0 since we are now moving to to another CC Type group
    sumTickerName = 0
    sumPriceDiff = 0
    sumPercDiff = 0
    sumTotalStockVolume = 0

    End If
    
    
Next rowline

'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'In this case, if the values in my 10th column are negative, then I want them to be red, otherwise they should be green.
'For the first row in column j to the last row in column j
For i = 2 To Range("J2").End(xlDown).row

'check to see if any values in cells are negative.
    If Cells(i, 10) < 0 Then
    'if they are, color them red
        Cells(i, 10).Interior.ColorIndex = 3
    'if they aren't, color them green.
    Else
        Cells(i, 10).Interior.ColorIndex = 4
        
    End If
    
Next i

        

End Sub
