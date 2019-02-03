Sub NextCells()
  Dim Stock_Name As String

  ' Set an initial variable for holding the total per stock
  Dim Volume_Total As Double
  Volume_Total = 0
  
  Dim Open_Price As Double
  Opent_Price = 0
  
  Dim Yearly_Change As Double
  Yearly_Change = 0
  
  Dim Percent_Change As Double
  Percent_Change = 0

  ' Keep track of the location for each stock in the summary table
  Dim Stock_Table_Row As Integer
  Stock_Table_Row = 2

  ' Set a variable for specifying the column of interest
  Dim column As Integer
  column = 1

  ' Loop through rows in the column
  For i = 2 To 43397
  
      Open_Price = Cells(i, 3).Value
      
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the Stock name
      Stock_Name = Cells(i, 1).Value
            
      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
      Yearly_Change = Cells(i, 6).Value - Open_Price
      Percent_Change = (Yearly_Change / Open_Price)

      ' Print the Credit Card Brand in the Stock Table
      Range("I" & Stock_Table_Row).Value = Stock_Name

      ' Print the Volumne Amount to the Stock Table
      Range("J" & Stock_Table_Row).Value = Volume_Total
      
      Range("K" & Stock_Table_Row).Value = Yearly_Change
      Range("L" & Stock_Table_Row).Value = Percent_Change

      ' Add one to the Stock Table Row
      Stock_Table_Row = Stock_Table_Row + 1
      
      ' Reset the Volume Total
      Volume_Total = 0
      Open_Price = 0
      Yearly_Change = 0
      Percent_Change = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
      
    End If
    
    Next i

End Sub