Attribute VB_Name = "Module1"
Sub Stocks()

Dim ws As Worksheet
Dim Ticker As String
Dim Ticker_Volume As Variant
Dim Summary_Table_Row As Integer
Dim YChange As Double
Dim PChange As Double
Dim StartRow As Long
Dim EndRow As Long
Dim MaxValue As Double
Dim MaxValueT As String
Dim MinValue As Double
Dim MinValueT As String
Dim HVolume As Double
Dim HVolumeT As String

For Each ws In Worksheets

ws.Select

    Ticker_Volume = 0
    Summary_Table_Row = 2
    StartRow = 2
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

        For I = 2 To LastRow
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                Ticker = Cells(I, 1).Value
                Ticker_Volume = Ticker_Volume + Cells(I, 7).Value
                YChange = (Cells(I, 6).Value) - (Range("C" & StartRow).Value)
                PChange = (YChange / (Range("C" & StartRow).Value))
                Range("I" & Summary_Table_Row).Value = Ticker
                Range("L" & Summary_Table_Row).Value = Ticker_Volume
                Range("J" & Summary_Table_Row).Value = YChange
                Range("K" & Summary_Table_Row).Value = PChange
                Summary_Table_Row = Summary_Table_Row + 1
                Ticker_Volume = 0
                StartRow = I + 1
            Else
                Ticker_Volume = Ticker_Volume + (Cells(I, 7).Value)
            End If
        Next I
    
    LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row
        
        For x = 2 To LastRow2
            If Cells(x, 10).Value > 0 Then
                Cells(x, 10).Interior.ColorIndex = 4
            ElseIf Cells(x, 10).Value < 0 Then
                Cells(x, 10).Interior.ColorIndex = 3
            End If
        Next x
    
    MaxValue = Cells(2, 11).Value
    MaxValueT = Cells(2, 9).Value
    MinValue = (Cells(2, 11).Value)
    MinValueT = Cells(2, 9).Value
    HVolume = Cells(2, 12).Value
    HVolumeT = Cells(2, 9).Value
    
    
        For j = 2 To LastRow2
            If Cells(j, 11).Value > MaxValue Then
                MaxValue = Cells(j, 11).Value
                MaxValueT = Cells(j, 9).Value
                Range("Q2").Value = MaxValue * 100 & "%"
                Range("P2").Value = MaxValueT
            End If
        Next j
            
        For k = 2 To LastRow2
            If Cells(k, 11).Value < MinValue Then
                MinValue = Cells(k, 11).Value
                MinValueT = Cells(k, 9).Value
                Range("Q3").Value = MinValue * 100 & "%"
                Range("P3").Value = MinValueT
            End If
        Next k
    
        For m = 2 To LastRow2
            If Cells(m, 12).Value > HVolume Then
                HVolume = Cells(m, 12).Value
                HVolumeT = Cells(m, 9).Value
                Range("Q4").Value = HVolume
                Range("P4").Value = HVolumeT
            End If
        Next m

Next ws

End Sub

