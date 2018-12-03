Sub Wall_Street()

  ' Set an initial variable for holding the Worksheet Count
  Dim WS_Count As Integer

  ' Set WS_Count equal to the number of worksheets in the active workbook.
  WS_Count = ActiveWorkbook.Worksheets.Count
  
  ' Begin the worksheet loop.
  For i = 1 To WS_Count
  
     ' Select each sheet individually
     sheets(i).Activate
    
     ' Determine the last row count
     Dim LastRow As Long
     LastRow = Cells(Rows.Count, 1).End(xlUp).Row
     
     ' Set an initial variable for holding the Ticker Symbol
     Dim Ticker_Symbol As String

     ' Set an initial variable for holding the Total Stock Volume
     Dim Stock_Volume As Double
     Stock_Volume = 0

     ' Keep track of the location for each Ticker Symbol in the summary table
     Dim Summary_Table_Row As Integer
     Summary_Table_Row = 2
     
     'Begin the row loop
     For j = 2 To LastRow

       ' Check if we are still within the same Ticker Symbol, if it is not...
       If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then

          ' Set the Ticker Symbol
          Ticker_Symbol = Cells(j, 1).Value
          
          ' Add to the Stock Volume
          Stock_Volume = Stock_Volume + Cells(j, 7).Value
      
          ' Print the Ticker Symbol in the Summary Table
          Range("I" & Summary_Table_Row).Value = Ticker_Symbol

          ' Print the Total Stock Volume to the Summary Table
          Range("L" & Summary_Table_Row).Value = Stock_Volume

          ' Add one to the Summary Table Row
          Summary_Table_Row = Summary_Table_Row + 1
      
          ' Reset the Stock Volume Total
          Stock_Volume = 0
      
          ' If the cell immediately following a row is the same Ticker Symbol...
        Else
          
           ' Add to the Stock Volume
           Stock_Volume = Stock_Volume + Cells(j, 7).Value

       'End check of Ticker Symbol
       End If
       
   ' Continue the table loop
    Next j
    
 ' Run individual applications for worksheet
   For Each sh In Worksheets
   
   Dim Yearly_Open_Price As Double
   Dim Yearly_Close_Price As Double
   Dim Yearly_Change As Double
   Dim Yearly_Max_Volume As Double
    
   Yearly_Close_Price = Cells(3, 6).Value
   Yearly_Open_Price = Cells(3, 3).Value
   Yearly_Change = Cells(3, 10).Value
   Cells(3, 10).Value = Cells(3, 6).Value - Cells(3, 3).Value

   ' Yearly_Max_Volume = Cells.Max(Range("L")).Value
   Yearly_Max_Volume = WorksheetFunction.Max(Range("L:L"))
   Range("Q4") = Yearly_Max_Volume
  
   Next sh
  
   ' Add one to the worksheet count
   WS_Count = WS_Count + 1
  
  ' Continue the worksheet loop
  MsgBox ("OK Calculation")
  
  Next i

End Sub
