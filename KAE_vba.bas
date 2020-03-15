Attribute VB_Name = "Module1"

Sub Macro10()
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"

  Dim ticker As String
  ticker = cells(i, 1).Value
  
  Dim Volume_total As Double
    stock_Name = 1
    Volume_total = 0

  ' Keep track in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
  For i = 2 To 6

    ' Check if we are still within the same name, if it is not...
    If cells(i + 1, stock_Name).Value <> cells(i, stock_Name).Value Then
    cells(i, 11).Value = stock_Name

      ' Set the Brand name
      stock_Name = cells(i, 1).Value

      ' Add to the Brand Total
    Volume_total = Volume_total + cells(i, 3).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("G" & Summary_Table_Row).Value = stock_Name

      ' Print the Brand Amount to the Summary Table
      Range("H" & Summary_Table_Row).Value = Volume_total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Volume_total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Volume_total = Volume_total + cells(i, 7).Value

    End If

  Next i

End Sub
    

Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = " \n14"

 Dim ticker As String
  ticker = cells(i, 1).Value


For i = 2 To

    ' Check if we are still within the same credit card brand, if it is not...
    If cells(i + 1, 1).Value <> cells(i, 1).Value Then
    
    cells(i + 1, 11).Value = cells(i, 1)

      ' Set the Brand name
      stock_Name = cells(i, 1).Value

    End If
  Next i
End Sub
Sub ticker()
Attribute ticker.VB_ProcData.VB_Invoke_Func = " \n14"

'define everything
Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim year_open As Double
Dim year_close_first As Double
Dim year_close_last As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim vol_total As Double
Dim Summary_Table_Row As Integer
Dim serial As Integer
Dim start As Long
start = 2


 'Set the initial value for the yearly_change to 0
  yearly_change = 0
  vol_total = 0


'this prevents my overflow error
On Error Resume Next

'run through each worksheet

For Each ws In ThisWorkbook.Worksheets
    
    'set headers
    ws.cells(1, 9).Value = "Ticker"
    ws.cells(1, 10).Value = "Yearly Change"
    ws.cells(1, 11).Value = "Percent Change"
    ws.cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Summary_Table_Row = 2


    'loop
        For i = 2 To ws.UsedRange.Rows.Count
            
             
             If cells(i + 1, 1).Value <> cells(i, 1).Value Then
            
            'find all the values
            ticker = cells(i, 1).Value
            vol = cells(i, 7).Value
            
            vol_total = vol_total + ws.cells(i, 7).Value
            If vol_total = 0 Then
            
            ws.cells(Summary_Table_Row, 9).Value = ticker
            ws.cells(Summary_Table_Row, 10).Value = 0
            ws.cells(Summary_Table_Row, 11).Value = 0
            ws.cells(Summary_Table_Row, 12).Value = 0
            
            Summary_Table_Row = Summary_Table_Row + 1
            Else
            'trying to find first non-zero starting value
            If cells(start, 3) = 0 Then
                For find_value = start To i
                If cells(find_value, 3).Value <> 0 Then
                start = find_value
                Exit For
            End If
                
            Next find_value
            End If
            
            yearly_change = (cells(i, 6) - cells(start, 3))
            percent_change = Round((yearly_change / cells(start, 3) * 100), 2)
            start = i + 1
            

            'insert values into summary
            ws.cells(Summary_Table_Row, 9).Value = ticker
            ws.cells(Summary_Table_Row, 10).Value = yearly_change
            ws.cells(Summary_Table_Row, 11).Value = percent_change
            ws.cells(Summary_Table_Row, 12).Value = vol_total
            Summary_Table_Row = Summary_Table_Row + 1
            
        End If
        vol_total = 0
        yearly_change = 0
         Else
            vol_total = vol_total + cells(i, 7)
        
    End If

'finish loop
    Next i
    
'PART 2----------------------------------------------


'assign max and min values

Dim greatest, least, greatest_tv As Long
Dim greatest_ticker As String
Dim least_ticker As String
    
    'set headers
    ws.cells(3, 13).Value = "Greatest Percent Increase"
    ws.cells(4, 13).Value = "Greatest Percent Decrease"
    ws.cells(4, 13).Value = "Greatest Total Volume"
    ws.cells(2, 14).Value = "Ticker"
    ws.cells(2, 15).Value = "Value"
    

        For j = 2 To ws.UsedRange.Rows.Count
        
        
        'find all the values
        ticker = ws.cells(i, 9).Value
        greatest = WorksheetFunction.Max(i, 11)
        least = WorksheetFunction.Min(i, 11)
        greatest_tv = WorksheetFunction.Max(i, 12)
       
        
          
    'insert values into summary
            ws.cells(3, 13).Value = greatest_ticker
            ws.cells(4, 13).Value = least_ticker
            ws.cells(5, 13).Value = greatest_tv
            ws.cells(3, 14).Value = greatest
            ws.cells(4, 14).Value = least
            ws.cells(5, 14).Value = vol_total
            
       Next j
        
    
ws.Columns("K").NumberFormat = "0.00%"

'Part 3 --------------------------------
'formatting

    'format columns colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g

'move to next worksheet
Next ws

End Sub
    
