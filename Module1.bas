Attribute VB_Name = "Module1"
Sub WallStreet()

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

    Dim WS_Name As String
    Dim Ticker As String
    Dim YrChange As Double
    Dim PctChange As Double
    Dim Volume As Double
    Volume = 0
    
    Dim Table_Row As Integer
    Table_Row = 2
    
    Dim Share_Open As Double
    Share_Open = Cells(2, 3).Value
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Columns("C").NumberFormat = "0.00"
    Columns("D").NumberFormat = "0.00"
    Columns("E").NumberFormat = "0.00"
    Columns("F").NumberFormat = "0.00"
    Columns("J").NumberFormat = "0.00"
    Columns("K").NumberFormat = "0.00%"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    WS_Name = ws.Name
    
    For x = 2 To LastRow
    
        If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
        
            Ticker = Cells(x, 1).Value
            YrChange = Cells(x, 6).Value - Share_Open
            
                If Share_Open = 0 Then
                PctChange = 0
                
                Else
                PctChange = (YrChange / Share_Open)
                
                End If
            
            Range("I" & Table_Row).Value = Ticker
            Range("J" & Table_Row).Value = YrChange
            
                If YrChange > 0 Then
                Range("J" & Table_Row).Interior.ColorIndex = 4
                
                ElseIf YrChange < 0 Then
                Range("J" & Table_Row).Interior.ColorIndex = 3
                
                End If
                
            Range("K" & Table_Row).Value = PctChange
            Volume = Volume + Cells(x, 7).Value
            Range("L" & Table_Row).Value = Volume
            Table_Row = Table_Row + 1
            Share_Open = Cells(x + 1, 3).Value
            Volume = 0
            
        Else
        
            Volume = Volume + Cells(x, 7).Value
            
        End If
        
    Next x
    
    Dim Tickr As String
    Dim HiPct As Double
    Dim LoPct As Double
    Dim HiVol As Double
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    HiPct = WorksheetFunction.Max(Range("K2:K" & LastRow))
    LoPct = WorksheetFunction.Min(Range("K2:K" & LastRow))
    HiVol = WorksheetFunction.Max(Range("L2:L" & LastRow))
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    Cells(2, 17).Value = HiPct
    Cells(3, 17).Value = LoPct
    Cells(4, 17).Value = HiVol
    
    For x = 2 To LastRow
    
        If Cells(x, 11).Value = HiPct Then
        
            Cells(2, 16).Value = Cells(x, 9).Value
        
        ElseIf Cells(x, 11).Value = LoPct Then
        
            Cells(3, 16).Value = Cells(x, 9).Value
            
        ElseIf Cells(x, 12).Value = HiVol Then
        
            Cells(4, 16).Value = Cells(x, 9).Value
            
        End If
        
    Next x
    
Next ws

End Sub

