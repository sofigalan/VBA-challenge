Attribute VB_Name = "Módulo1"
Sub StockData():
    
    For Each ws In Worksheets

        Dim Ticker As String
        Dim totalStockVolume As Double
        
        totalStockVolume = 0
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percentage Change"
        ws.Range("L1") = "Total Stock Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
        
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = totalStockVolume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                totalStockVolume = 0
            Else
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i

    Next ws
    
End Sub

Sub yearlyChange():

    For Each ws In Worksheets

        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim openYear As Double
        Dim closeYear As Double
        
        openYear = 0
        closeYear = 0
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 1
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                openYear = ws.Cells(i, 3).Value
                
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value Then
            
                closeYear = ws.Cells(i, 6).Value
                
                Summary_Table_Row = Summary_Table_Row + 1
                yearlyChange = 0
                percentChange = 0
                
                yearlyChange = closeYear - openYear
                ws.Range("J" & Summary_Table_Row).Value = yearlyChange
                If openYear = 0 Then
                    percentChange = 1
                Else
                    percentChange = ((closeYear / openYear) - 1)
                End If
                ws.Range("K" & Summary_Table_Row).Value = percentChange
            Else
            End If
        Next i
    
    Next ws
    
End Sub

Sub format():

    For Each ws In Worksheets
    
        Dim lastrow As Integer
                
        lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            End If
        Next i
    
        ws.Range("K:K").NumberFormat = "0.00%"

    Next ws
    
End Sub

