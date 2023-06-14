Sub Challenge2()
'Define Variables

Dim Ticker As String
Dim i As Double
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

'Define column in which we'll find the ticker value

Dim Column As Integer
Dim RowsDone As Double


'Set values
Column = 1
RowsDone = 1   'to keep track of the rows already processed
Stock_Volume = 0


'Define last row of data
lastrow = Cells(Rows.Count, 1).End(xlUp).Row 'as discussed in class.


'Begin


For i = 2 To lastrow

    Stock_Volume = Stock_Volume + Cells(i, 7)
    
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then     'so the row after doesn't match this row. As used in credit charges exercise in class.
       
       'print name of ticker to summary table
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
                    Cells(Summary_Table_Row, 13).Value = Format(0, "#.##%") 'This format function derived from information from 'davidjaimes' see README file for URL.
            Else
                    Cells(Summary_Table_Row, 13).Value = Format((closePrice / openPrice) - 1, "#.##%")
            End If
            
        'print to table
        Range("L" & Summary_Table_Row).Value = Yearly_Change
                        
        'print total volume to summary table
        Range("N" & Summary_Table_Row).Value = Stock_Volume
        
        'Reset volume total
        Stock_Volume = 0
        
        'Add a row to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
    End If
    
Next i

'create analysis table

Cells(1, 19).Value = "Ticker"
Cells(1, 20).Value = "Value"
Cells(2, 18).Value = "Greatest % increase"
Cells(3, 18).Value = "Greatest % decrease"
Cells(4, 18).Value = "Greatest Total Volume"

'Definitions:

Dim PricePercents As Range
Set PricePercents = Range(Cells(2, 13), Cells(Summary_Table_Row - 1, 13))
Dim TSV As Range
Set TSV = Range(Cells(2, 14), Cells(Summary_Table_Row - 1, 14))

Cells(2, 20).Value = Format(Application.WorksheetFunction.Max(PricePercents), "#.##%")    'From information from https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.max
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
'Format

Columns(12).EntireColumn.AutoFit
Columns(14).EntireColumn.AutoFit
Columns(18).EntireColumn.AutoFit
Columns(20).EntireColumn.AutoFit
Rows(1).Font.Bold = True



End Sub


