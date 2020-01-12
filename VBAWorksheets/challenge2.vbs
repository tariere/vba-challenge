Attribute VB_Name = "Module3"


Sub runallsheets()

'create a variable for the worksheet
Dim sheet As Worksheet

'create a variable for my loop
Dim i As Integer
        'A for loop that will go through each sheet in the worksheets in the spreadsheet
         For Each sheet In Worksheets
          
          'This will select the sheet
          sheet.Select
         
         'this will run my macro
         Call wsExtrema

    
        'this will move to the next sheet in the work book
         Next
End Sub
