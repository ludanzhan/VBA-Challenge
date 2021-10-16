Attribute VB_Name = "Module1"
Sub tickersummary()
        For Each ws In Worksheets
            ws.Range("J1").Value = "Ticker"
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L1").Value = "Percent Change"
            ws.Range("M1").Value = "Total"
            
            Dim ticker As String
            Dim Totalval As Double
            Dim yearClose As Double
            Dim yearbegin As Double
            Dim yearChange As Double
            Dim percentChange As Double
            
            Dim summaryTableRow As Integer
            Dim Change  As Integer
            
            summaryTableRow = 2
            
            
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Change = Application.WorksheetFunction.CountIf(ws.Range("A1:A" & lastRow), ws.Range("A2")) - 1
      
            
            Totalval = 0

            For i = 2 To lastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                     ticker = ws.Cells(i, 1).Value
                     Totalval = Totalval + ws.Cells(i, 7).Value
                     
                     yearClose = ws.Cells(i, 6).Value
                     yearbegin = ws.Cells(i - Change, 3).Value
                     yearChange = yearClose - yearbegin
                    ' percentChange = yearChange / yearbegin
                
                      ws.Range("J" & summaryTableRow).Value = ticker
                      ws.Range("M" & summaryTableRow).Value = Totalval
                      ws.Range("K" & summaryTableRow).Value = yearChange
                      ' ws.Range("L" & summaryTableRow).Value = percentChange
                      ws.Range("N" & summaryTableRow).Value = yearbegin
                      
                      summaryTableRow = summaryTableRow + 1
                Totalval = 0
                Else
                      ticker = ws.Cells(i, 1).Value
                      Totalval = Totalval + ws.Cells(i, 7).Value
                End If
            Next i
        Next ws
End Sub

