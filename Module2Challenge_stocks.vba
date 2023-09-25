Sub change_worksheet():
Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Call stocks
    Next

End Sub


Sub stocks():
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Dim current As String
    Dim i, counter_list, color As Integer
    Dim op, cl, yc, pc, total, max As Double
    
    counter_list = 2
    total = 0
    op = Cells(2, 3).Value
    
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        current = Cells(i, 1)
        total = total + Cells(i, 7).Value
        
        If current <> Cells(i + 1, 1).Value Then
            Cells(counter_list, 9).Value = current
            cl = Cells(i, 6).Value
            yc = cl - op
            pc = yc / op
            
            Cells(counter_list, 10).Value = yc
            If yc < 0 Then
                color = 3
            Else
                color = 4
            End If
            
            Cells(counter_list, 10).Interior.ColorIndex = color
            Cells(counter_list, 11).Value = pc
            Cells(counter_list, 11).NumberFormat = "0.00%"
            Cells(counter_list, 12).Value = total
            
            op = Cells(i + 1, 3).Value
            total = 0
            counter_list = counter_list + 1
        End If
    Next i
    
    Dim cp, ctv, gi, gd, gtv As Double
    Dim ct, tgi, tgd, tgtv As String
    
    
    gi = Cells(2, 11).Value
    gd = Cells(2, 11).Value
    gtv = Cells(2, 12).Value
    
    For i = 2 To Range("I:I").SpecialCells(xlCellTypeConstants).Count
        cp = Cells(i, 11).Value
        ctv = Cells(i, 12).Value
        If cp > gi Then
        gi = cp
        tgi = Cells(i, 9).Value
        End If
        
        If cp < gd Then
        gd = cp
        tgd = Cells(i, 9).Value
        End If
        
        If ctv > gtv Then
        gtv = ctv
        tgtv = Cells(i, 9).Value
        End If
    Next i
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 16).Value = tgi
    Cells(3, 16).Value = tgd
    Cells(4, 16).Value = tgtv
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    Cells(2, 17).Value = gi
    Cells(3, 17).Value = gd
    Cells(4, 17).Value = gtv
    
    
    
End Sub