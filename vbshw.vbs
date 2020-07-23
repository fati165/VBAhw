Attribute VB_Name = "Module1"
Sub stock()
'create a loop for stocks in a given year

'create/set variables
Dim workS As Worksheet
Dim tick As String
Dim yChange As Double
Dim pChange As Double
Dim totalStock As Double
Dim lastRow As Long
Dim openStock As Double
Dim OutputRow As Double
Dim closeStock As Double

'create for loop, within for loop nest the if statements
For Each workS In Worksheets

    'create/ assign the new column labels
    workS.Cells(1, 9).Value = "Ticker"
    workS.Cells(1, 10).Value = "Yearly Change"
    workS.Cells(1, 11).Value = "Percent Change"
    workS.Cells(1, 12).Value = "Total Stock Volume"

 
    'initialize the variables we use to calculate
    lastRow = workS.Cells(Rows.Count, 1).End(xlUp).Row
    openStock = workS.Cells(2, 3).Value
    tick = workS.Cells(2, 1).Value
    OutputRow = 2
    totalStock = 0
    yChange = workS.Cells(2, 10).Value
    yChange = 0
'ychange is closestock - openstock

'for loop, nest if statements
For i = 2 To lastRow ' means we start output row on 2

    If (tick <> workS.Cells(i, 1).Value) Then
            closeStock = workS.Cells(i - 1, 6).Value '-1 means the equal value but stepping back to grab that value
            
            'yearly change calculate
            workS.Cells(OutputRow, 10).Value = closeStock - openStock
            workS.Cells(OutputRow, 9).Value = tick
            
            'highlight positive or negative change??
            If yChange >= 0 Then
            
                workS.Cells(OutputRow, 10).Interior.ColorIndex = 4
                
            Else
            
                workS.Cells(OutputRow, 10).Interior.ColorIndex = 3
                    
            End If
            
            'if inside
            'percent change calculate
            ' opening stock not equal to zero, then calculate. otherwise don't calculate and just make it zero
            If openStock <> 0 Then
                pChange = (closeStock - openStock) / openStock
            
            Else
                pChange = 0 'zero values don't need to calculate
                
            End If
            
            'output yearly change
            'change pchange values to percent
            workS.Cells(OutputRow, 11).Value = pChange
            workS.Cells(OutputRow, 11).NumberFormat = "0.00%"

            'output total volume/ stock
            workS.Cells(OutputRow, 12).Value = totalStock
            
            'reset to start over
            tick = workS.Cells(i, 1).Value
            openStock = workS.Cells(i, 3).Value
            OutputRow = OutputRow + 1
            totalStock = workS.Cells(i, 7).Value
            
            Else
                totalStock = totalStock + workS.Cells(i, 7).Value
                
        
    End If
    Next i


Next


End Sub
