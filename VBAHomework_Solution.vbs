Attribute VB_Name = "Module1"
Sub Stock()

Dim Total As Double
Dim i, j As Long
Dim LastRow As Double
Dim YearlyChange As Double, Percentchange As Double
Dim tickersym As String
Dim OpenValue As Double, LastValue As Double
Dim sheet As Worksheet

For Each sheet In ThisWorkbook.Worksheets

LastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
'define last row in column A
j = 2
OpenValue = sheet.Cells(2, 3).Value
'Set the initial open value
'FOr loop to start going thru rows

For i = 2 To LastRow
    Total = Total + sheet.Cells(i, 7).Value
    'cal total as row counting
   
    
    If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
        tickersym = sheet.Cells(i, 1).Value
        sheet.Range("I" & j).Value = tickersym
        sheet.Range("L" & j).Value = Total
        'output ticker symbol and total when next ticker is different than prious one
        
        LastValue = sheet.Cells(i, 6).Value
        YearlyChange = (LastValue - OpenValue)
        sheet.Range("J" & j).Value = YearlyChange
        'Output change into column J
        If YearlyChange < 0 Then
        sheet.Range("J" & j).Interior.ColorIndex = 3
        Else:
        sheet.Range("J" & j).Interior.ColorIndex = 4
        End If
        
        If (OpenValue = 0 And LastValue = 0) Then
            Percentchange = 0
        ElseIf (OpenValue = 0 And LastValue <> 0) Then
            Percentchange = -1
           
        Else:
            Percentchange = (YearlyChange / OpenValue)
           
        End If
        sheet.Range("K" & j).Value = Percentchange
        sheet.Range("K" & j).NumberFormat = "0.00%"
            'Output %change into column K
        Total = 0
        OpenValue = sheet.Cells(i + 1, 3).Value
        j = j + 1
    End If
    
    'If i = LastRow - 1 Then
        'Range("L" & j).Value = Total
        'tickersym = Cells(LastRow, 1).Value
       ' Range("I" & j).Value = tickersym
        'LastValue = Cells(LastRow, 6).Value
       ' YearlyChange = (LastValue - OpenValue)
       ' Range("J" & j).Value = YearlyChange
        'If YearlyChange < 0 Then
        'Range("J" & j).Interior.ColorIndex = 3
       ' Else:
       ' Range("J" & j).Interior.ColorIndex = 4
       ' End If
        
       ' If (OpenValue = 0 And LastValue = 0) Then
           ' Percentchange = 0
        'ElseIf (OpenValue = 0 And LastValue <> 0) Then
            'Percentchange = -1
            
        'Else:
            'Percentchange = (YearlyChange / OpenValue)
            'Range("K" & j).Value = Percentchange
            'Range("K" & j).NumberFormat = "0.00%"
            
       ' End If
    'End If
   
Next i
sheet.Range("O2") = "%" & Application.WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
sheet.Range("O3") = "%" & Application.WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
sheet.Range("O4") = Application.WorksheetFunction.Max(Range("L2:L" & LastRow))
Next sheet
End Sub

