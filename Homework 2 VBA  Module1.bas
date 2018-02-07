Attribute VB_Name = "Module1"
Sub forEachWs()

For Each ws In ActiveWorkbook.Worksheets
Call stock_analysis(ws)
Next

End Sub


Sub stock_analysis(ws)

    Dim ticker As String

    Dim total_volume As Double
    total_volume = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

         
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                Cells(1, 9).Value = "Ticker"
                Cells(1, 10).Value = "Total Volume"
        

                ticker = Cells(i, 1).Value

                total_volume = total_volume + Cells(i, 3).Value

                Range("I" & Summary_Table_Row).Value = ticker

                Range("J" & Summary_Table_Row).Value = total_volume

                Summary_Table_Row = Summary_Table_Row + 1
      
                total_volume = 0

            Else

                total_volume = total_volume + Cells(i, 3).Value

        End If

        Next i

End Sub

