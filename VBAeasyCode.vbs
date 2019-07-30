Sub TickerTotals()
   Dim WS_Count As Integer
   Dim S As Integer
        
   WS_Count = ActiveWorkbook.Worksheets.Count
         
    For S = 1 To WS_Count
         ActiveWorkbook.Worksheets(S).Select
        Dim I As Integer
        Dim e As Integer
        Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I2"), Unique:=True
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Total Stock Volume"
        

        e = Cells(Rows.Count, 9).End(xlUp).Row

        For I = 2 To e

            Cells(I, 10).Value = Application.SumIf(Range("A:A"), "=" & Cells(I, 9).Value, Range("G:G"))

        Next I
    Next S

End Sub
