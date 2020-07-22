Attribute VB_Name = "Module1"
Sub Stock_Market_Analyst()
    Dim Current As Worksheet
    'On Error Resume Next
    For Each Current In ThisWorkbook.Worksheets
        Dim i As Long
        Dim Ticker As String
        Dim Ticker_Row As Long
        Dim Yearly_Change As Double
        Dim Yearly_Change_Row As Long
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Percent_Change As Double
        Dim Percent_Change_Row As Long
        Dim Stock_Volume As Double
        Dim Sum_Stock_Volume As Double
        Dim Stock_Volume_Row As Long
        Dim Percent_Increase As Double
        Ticker_Row = 2
        Yearly_Change_Row = 2
        Percent_Change_Row = 2
        Stock_Volume_Row = 2
        Sum_Stock_Volume = 0
        For i = 2 To Current.UsedRange.Rows.Count
            Ticker = Current.Cells(i, 1).Value
            Opening_Price = Current.Cells(i, 3).Value
            Closing_Price = Current.Cells(i, 6).Value
            Stock_Volume = Current.Cells(i, 7).Value
            If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
                Current.Range("I" & Ticker_Row).Value = Ticker
                Ticker_Row = Ticker_Row + 1
                Yearly_Change = Closing_Price - Opening_Price
                Current.Range("J" & Yearly_Change_Row).Value = Yearly_Change
                Yearly_Change_Row = Yearly_Change_Row + 1
                
                If Opening_Price = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / Opening_Price
                End If
                Current.Range("K" & Percent_Change_Row).Value = Percent_Change
                Current.Range("K" & Percent_Change_Row).NumberFormat = "###,###0.0#%"
                Percent_Change_Row = Percent_Change_Row + 1
                'Yearly_Change = 0
                Sum_Stock_Volume = Sum_Stock_Volume + Stock_Volume
                Current.Range("L" & Stock_Volume_Row).Value = Sum_Stock_Volume
                Stock_Volume_Row = Stock_Volume_Row + 1
                Sum_Stock_Volume = 0
            Else
                Sum_Stock_Volume = Sum_Stock_Volume + Stock_Volume
            End If
            'Current.Cells(i, 10).ClearFormats
            If Current.Cells(i, 10).Value > 0 Then
                Current.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Current.Cells(i, 10).Value < 0 Then
                Current.Cells(i, 10).Interior.ColorIndex = 3
            End If
            ' For j = 2 To Current.UsedRange.Rows.Count
               ' If Yearly_Change > current.cells(j,10).value then
               '
        Next i
        'Current.Range("L" & Stock_Volume_Row).Value = Sum_Stock_Volume
    Next Current
End Sub
