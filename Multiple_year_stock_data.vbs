Attribute VB_Name = "Module1"
Sub StockData():
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Columns.AutoFit
    ws.Activate
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        Dim Ticker_Name As String
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Annual_Change As Double
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        Opening_Price = Cells(2, Column + 2).Value
        
        For i = 2 To LastRow
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                
                Closing_Price = Cells(i, Column + 5).Value
                
                Annual_Change = Closing_Price - Opening_Price
                Cells(Row, Column + 9).Value = Annual_Change
                
                If (Opening_Price = 0 And Closing_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Annual_Change / Opening_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                Row = Row + 1
                               
                Opening_Price = Cells(i + 1, Column + 2)
                
                Volume = 0
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        YCLastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
            
        Next j
        
    Next ws
    
End Sub
