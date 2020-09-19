Sub VBAChallenge()

    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
    Next ws

End Sub


Sub CalculateSummary()

'Ticker Name
Dim Ticker As String

'Total Volume
Dim Volume As Double
Volume = 0

'Opening Price
Dim OPrice As Double
OPrice = Cells(2, 3).Value

'Closing Price
Dim CPrice As Double

'Yearly Change
Dim YearC As Double

Dim PerC As Double

'Summary Table
Dim SumRow As Integer
SumRow = 2

'Number of Rows
Dim NRow As Long
NRow = Cells(Rows.Count, "A").End(xlUp).Row

For CRow = 2 To NRow

    If Cells(CRow + 1, 1).Value <> Cells(CRow, 1).Value Then
    
        'Setting Ticker Name
        Ticker = Cells(CRow, 1).Value
        
        'Ticker Name in Summary Table
        Range("I" & SumRow).Value = Ticker
        
        'Add Volume
        Volume = Volume + Cells(CRow, 7).Value
        
        'Volume in Summary Table
        Range("L" & SumRow).Value = Volume
        
        'Closing Price
        CPrice = Cells(CRow, 6).Value
        
        'Calculate Yearly Change
        YearC = (CPrice - OPrice)
        
        'Yearly Change in Summary Table
        Range("J" & SumRow).Value = YearC
        
        'Calculate Percent Change with Divisibility Checker
        If OPrice = 0 Then
            
            PerC = 0
        Else
            
            PerC = (YearC / OPrice) * 100
        
        End If
        
        'Percent Change in Summary Table
        Range("K" & SumRow).Value = PerC
        
        'Adding Row to Summary Table
        SumRow = SumRow + 1
    
        'Reset Volume
        Volume = 0
        
        'Reset Opening Price
        OPrice = Cells(CRow + 1, 3)
        
    Else
    
        'Add Volume
        Volume = Volume + Cells(CRow, 7).Value
        
    End If
    
Next CRow

'Last Row in Summary Table
LRow_Sum = Cells(Rows.Count, 10).End(xlUp).Row

'Color Code Yearly Change
For CRow = 2 To LRow_Sum

    If Cells(CRow, 10).Value > 0 Then
        
        'Positive Yearly Change = Green
        Cells(CRow, 10).Interior.ColorIndex = 4
        
    Else
    
        'Negative Yearly Change = Red
        Cells(CRow, 10).Interior.ColorIndex = 3
        
    End If

Next CRow
    
End Sub


Sub SetTitle()

Range("I:Q").Value = ""
Range("I:Q").Interior.ColorIndex = 0

' Set title row
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


End Sub