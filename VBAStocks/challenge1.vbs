Attribute VB_Name = "Module2"
Sub wsExtrema()

'First, I want my i intial script to run and to provide all of my summarized values and place the values where they belong.
Call wallstreetsummaries

'Then, I'll create some rows and columns to capture where I want to place my extrema values
'Creating some Columns to store the values.
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Creating some rows to label the values
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'I'm going to do a little formatting of my cells to make sure that the data is more readable
Range("O1:Q4").Font.Bold = True
Range("O1:Q4").Font.Italic = True
Columns("O:Q").AutoFit

'Next, I'll create some variables to hold my values
'This variable is for the max percentage increase
Dim grtPercInc As Double

'This variable is for the min percentage decrease
Dim grtPercDec As Double

'This variable is for the max total volume
Dim totalVolInc As Double

'Now that I have these variables, I want to loop through my summarized values to find the specific items i'm looking for.
'First, I'll set up a for loop

For i = 2 To Range("K2").End(xlDown).row

'Then I'll set up a way to populate my variables.  There is a function called worksheet.function.max and worksheet.function.min that I'll use here

    'This will get the max percentage increase using the function and put in in my variable.
    grtPercInc = WorksheetFunction.Max(Range("K:K"))
    
    'This will get the max percentage increase using the function and put in in my variable.
    grtPercDec = WorksheetFunction.Min(Range("K:K"))
    
    'This will get the max percentage increase using the function and put in in my variable.
    totalVolInc = WorksheetFunction.Max(Range("L:L"))
    
    'I'll set up if statements to see if any values in my summary list match the maxes i found. If so, then I'll take the ticker name and the value and place it in the section I created for this analysis
    
        If Cells(i, 11).Value = grtPercInc Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
            
        ElseIf Cells(i, 11).Value = grtPercDec Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
        
        ElseIf Cells(i, 12).Value = totalVolInc Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
        End If
    
    'update the formatting to look like percentages
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
Next i


End Sub
