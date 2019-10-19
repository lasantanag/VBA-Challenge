Sub applytoall()

Dim xSh As Worksheet
Application.ScreenUpdating = False
For Each xSh In Worksheets

xSh.Select

Call stock_volume_counter

Next
Application.ScreenUpdating = True



End Sub


Sub stock_volume_counter()

  ' Set an initial variable for holding the stock name
  Dim stock As String
  Dim counter As Integer
  
  
  ' Set an initial variable for holding the total per stock brand
  Dim Stock_Total As Double
  Stock_Total = 0
  
    'Set initial variable for Yearly_Open, Yearly_Close and Yearly_Change
    Dim Yearly_Open As Double
    Dim Yearly_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    
    
    Yearly_Open = 0
    Yearly_Close = 0
    Yearly_Change = 0
    Percent_Change = 0

  
  ' Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Titles
  Range("H1").Value = "Ticker Name"
  Range("I1").Value = "Total Stock Value"
  Range("J1").Value = "Yearly_Change"
  Range("K1").Value = "Yearly_Open"
  Range("L1").Value = "Yearly_Close"
  Range("M1").Value = "Percent_Change"
  

'allocation and calculation of Yearly_Open
  
  'counter creation to allocate stock name for printing
    counter = 2
    
    
    ' Loop through all stocks
  For i = 1 To 43398
    
    ' Check if we are still within the same stock brand, if we are not...
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
       'Allocate Yearly_Open
        Yearly_Open = Cells(i + 1, 3).Value
    
        'Print Yearly_Open
        Cells(counter, 11) = Yearly_Open
        
        'Reset Yearly_Open
        Yearly_Open = 0
    
    
        'Independer counter to allocate unique stock name one after another
        counter = counter + 1
    
        End If
    
    Next i
    
    
    'allocation and calculation of Stock, Stock_Total, Yearly_Close, Percent_Change
    
    'counter creation to allocate stock name for printing
    counter = 2
    

  ' Loop through all stocks
  For i = 1 To 43398
    
    
       ' Check if we are still within the same stock brand, if we are not...
        If Cells(i + 1, 1).Value <> Cells(i + 2, 1).Value Then
    
        stock = Cells(i + 1, 1).Value
       
      ' put name in cells (H2) and below the unique stock names
      Cells(counter, 8).Value = stock
      
      ' Add to the Stock_Total this last value before Else
      Stock_Total = Stock_Total + Cells(i + 1, 7).Value
      
      ' Print the Stock_Total
      Cells(counter, 9).Value = Stock_Total

      
        'Independer counter to allocate unique stock name one after another
     counter = counter + 1
     
      ' Reset the Stock_Total
      Stock_Total = 0
      
      
      ' Reset Yearly_Close
       Yearly_Close = 0
       
      ' Reset Yearly_Change
       'Yearly_Change = 0
       
    
    'Allocate close of given stock in dec 31
    Yearly_Close = Cells(i + 1, 6).Value
    
    'Calculation of Yearly_Change
    Yearly_Change = Yearly_Close - Yearly_Open
    
    'Print
    Cells(counter - 1, 10) = Yearly_Change
    Cells(counter - 1, 12) = Yearly_Close


    ' If the cell immediately following a row is the same brand...
    
    Else

      ' Add to the Stock_Total
    Stock_Total = Stock_Total + Cells(i + 1, 7).Value

    End If

  Next i
  
  
  lastrow = Cells(Rows.Count, 11).End(xlUp).Row
  
  For i = 2 To lastrow
  
    Cells(i, 10).Value = Cells(i, 12).Value - Cells(i, 11).Value
    
    Next i
    
    
    For i = 2 To lastrow
    'Calculation of Percent_Change
    Percent_Change = ((Cells(i, 12).Value) - (Cells(i, 11).Value)) / (Cells(i, 11).Value)
        
    'Print
    Cells(i, 13) = Percent_Change
    
    Next i
    
    
    For i = 2 To lastrow
    
    If Cells(i, 10).Value < 0 Then
    
    Range("J" & i).Interior.ColorIndex = 3
    
  
  Else
    
    Range("J" & i).Interior.ColorIndex = 4
    
    End If
    
    
    Next i
    
   
End Sub