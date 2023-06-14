Sub Challenge2_allsheets()

   ' Loop through all sheets
    For Each ws In Worksheets

        ws.Activate
        
        YearlyChange

    Next ws





End Sub



Sub YearlyChange()
'Define Variables

Dim Ticker As String
Dim i As Long
Dim lastrow As Double
Dim Stock_Volume As Double
Dim Yearly_Change As Double
Dim Percentage_Change As Double


'summary table info
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Cells(1, 11).Value = "Ticker"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "% Change"
Cells(1, 14).Value = "Total Stock Volume"

'Define column
Dim Column As Integer
Dim RowsDone As LongLong


'Set values
Column = 1
RowsDone = 1
Stock_Volume = 0



lastrow = Cells(Rows.Count, 1).End(xlUp).Row





For i = 2 To lastrow

    Stock_Volume = Stock_Volume + Cells(i, 7)
    
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
       
      
        Range("K" & Summary_Table_Row).Value = Cells(i, 1).Value
        openPrice = Cells(RowsDone + 1, 3).Value
        closePrice = Cells(i, 6).Value
        Yearly_Change = closePrice - openPrice
        Cells(i, 12).Value = Yearly_Change
        RowsDone = i

        
            
            If Yearly_Change < 0 Then
                    Cells(Summary_Table_Row, 12).Interior.ColorIndex = 3
            ElseIf Yearly_Change >= 0 Then
                    Cells(Summary_Table_Row, 12).Interior.ColorIndex = 4
            End If
                    
            If openPrice = 0 Then
                    Cells(Summary_Table_Row, 13).Value = Format(0, "#.##%")
            Else
                    Cells(Summary_Table_Row, 13).Value = Format((closePrice / openPrice) - 1, "#.##%")
            End If
            
        'print to table
        Range("L" & Summary_Table_Row).Value = Yearly_Change
                        
      
        Range("N" & Summary_Table_Row).Value = Stock_Volume
        
       
        Stock_Volume = 0
        
      
        Summary_Table_Row = Summary_Table_Row + 1
        
    End If
    
Next i

Cells(1, 19).Value = "Ticker"
Cells(1, 20).Value = "Value"
Cells(2, 18).Value = "Greatest % increase"
Cells(3, 18).Value = "Greatest % decrease"
Cells(4, 18).Value = "Greatest Total Volume"

Dim PricePercents As Range
Set PricePercents = Range(Cells(2, 13), Cells(Summary_Table_Row - 1, 13))
Dim TSV As Range
Set TSV = Range(Cells(2, 14), Cells(Summary_Table_Row - 1, 14))

Cells(2, 20).Value = Format(Application.WorksheetFunction.Max(PricePercents), "#.##%")
Cells(3, 20).Value = Format(Application.WorksheetFunction.Min(PricePercents), "#.##%")
Cells(4, 20).Value = Application.WorksheetFunction.Max(TSV)

For i = 2 To Summary_Table_Row - 1
    
    If Cells(i, 13) = Cells(2, 20) Then
        Cells(2, 19).Value = Cells(i, 11).Value
    End If
    
    If Cells(i, 13) = Cells(3, 20) Then
        Cells(3, 19).Value = Cells(i, 11).Value
    End If
    
        If Cells(i, 14) = Cells(4, 20) Then
        Cells(4, 19).Value = Cells(i, 11).Value
    End If


Next i

Columns(14).EntireColumn.AutoFit
Columns(18).EntireColumn.AutoFit
Columns(20).EntireColumn.AutoFit


End Sub


