Attribute VB_Name = "Module1"
Sub lkjh()

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
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Summary_Table_Row = 2


    'loop
        For i = 2 To ws.UsedRange.Rows.Count
            
             
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'find all the values
            ticker = Cells(i, 1).Value
            vol = Cells(i, 7).Value
            
            vol_total = vol_total + ws.Cells(i, 7).Value
            If vol_total = 0 Then
            
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = 0
            ws.Cells(Summary_Table_Row, 11).Value = 0
            ws.Cells(Summary_Table_Row, 12).Value = 0
            
            Summary_Table_Row = Summary_Table_Row + 1
            Else
            'trying to find first non-zero starting value
            If Cells(start, 3) = 0 Then
                For find_value = start To i
                If Cells(find_value, 3).Value <> 0 Then
                start = find_value
                Exit For
            End If
                
            Next find_value
            End If
            
            yearly_change = (Cells(i, 6) - Cells(start, 3))
            percent_change = Round((yearly_change / Cells(start, 3) * 100), 2)
            start = i + 1
            

            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol_total
            Summary_Table_Row = Summary_Table_Row + 1
            
        End If
        vol_total = 0
        yearly_change = 0
         Else
            vol_total = vol_total + Cells(i, 7)
        
    End If

'finish loop
    Next i
    
ws.Columns("K").NumberFormat = "0.00%"

'Part 3 --------------------------------
'formatting

    'format columns colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
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

