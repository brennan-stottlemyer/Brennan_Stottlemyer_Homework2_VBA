Attribute VB_Name = "Module1"
Sub StockData():
    Dim ws As Worksheet
    Dim i As Long
    Dim Ticker As Integer
    Dim VolumeTotal As LongLong
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    For Each ws In ActiveWorkbook.Worksheets
    MsgBox (ws.Name)
    
        i = 2
        Ticker = 1
        VolumeTotal = 0
        OpenPrice = 0
        ClosePrice = 0
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        
        Do Until IsEmpty(ws.Cells(i, 1))

            ' If the cell before doesn't match the current cell, calculate open price
            If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
                OpenPrice = ws.Cells(i, 3)

            End If
            
            ' If the cell after doesn't match the current cell
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                Ticker = Ticker + 1
                'Record Name of Ticker
                ws.Cells(Ticker, 9) = ws.Cells(i, 1)

                'Calculating Yearly Change using Open Price and Close Price
                ClosePrice = ws.Cells(i, 6)

                ws.Cells(Ticker, 10) = ClosePrice - OpenPrice

                If OpenPrice = 0 Then
                    ws.Cells(Ticker, 11) = 0

                Else:
                    'Calculating Percent Change
                    ws.Cells(Ticker, 11) = (ws.Cells(Ticker, 10)) / OpenPrice

                End If
                
                ' Calculating Sum of Volume
                VolumeTotal = VolumeTotal + ws.Cells(i, 7)
                ws.Cells(Ticker, 12) = VolumeTotal
                VolumeTotal = 0
                
                'Color Yearly Change cell green if positive change, red if negative
                If ws.Cells(Ticker, 10) >= 0 Then
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 4

                Else:
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 3

                End If

            'If cell after matches current sell, continue to add up sum of volume
            Else:
                VolumeTotal = VolumeTotal + ws.Cells(i, 7)
                
            End If
            
            i = i + 1
            
        Loop
        
    Next ws

End Sub

