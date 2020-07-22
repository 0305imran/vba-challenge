Sub Stock_Market_Analyst()
    Dim Current As Worksheet
    
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
    
        Ticker_Row = 2
        Yearly_Change_Row = 2
        Percent_Change_Row = 2
        Stock_Volume_Row = 2
        Sum_Stock_Volume = 0
          
        For i = 2 To Current.UsedRange.Rows.Count
            Opening_Price = Current.Cells(Ticker_Row, 3).Value
            Closing_Price = Current.Cells(i, 6).Value
            Stock_Volume = Current.Cells(i, 7).Value
            Yearly_Change = Opening_Price
            
            If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
                Ticker = Current.Cells(i, 1).Value
                Current.Range("I" & Ticker_Row).Value = Ticker
                Ticker_Row = i + 1
        
                Yearly_Change = Closing_Price - Opening_Price
                Current.Range("J" & Yearly_Change_Row).Value = Yearly_Change
                
                If Current.Cells(Yearly_Change_Row, 10).Value > 0 Then
                    Current.Cells(Yearly_Change_Row, 10).Interior.ColorIndex = 4

                ElseIf Current.Cells(Yearly_Change_Row, 10).Value < 0 Then
                    Current.Cells(Yearly_Change_Row, 10).Interior.ColorIndex = 3
            
                End If
                
                Yearly_Change_Row = Yearly_Change_Row + 1
                
                If Opening_Price = 0 Then
                    For j = Ticker_Row To i
                    
                        If Current.Cells(j, 3).Value <> 0 Then
                            Opening_Price = Current.Cells(Ticker_Row, 3).Value
                            Exit For
                            
                        End If
                        
                    Next j
                    
                    Percent_Change = 0
                    
                Else
                    Percent_Change = Yearly_Change / Opening_Price
                    
                End If
                
                Current.Range("K" & Percent_Change_Row).Value = Percent_Change
                Current.Range("K" & Percent_Change_Row).NumberFormat = "###,###0.0#%"
                Percent_Change_Row = Percent_Change_Row + 1
                
                Sum_Stock_Volume = Sum_Stock_Volume + Stock_Volume
                Current.Range("L" & Stock_Volume_Row).Value = Sum_Stock_Volume
                Stock_Volume_Row = Stock_Volume_Row + 1
                Sum_Stock_Volume = 0
                
            Else
                Sum_Stock_Volume = Sum_Stock_Volume + Stock_Volume
                    
            End If
              
        Next i
        
    Next Current

End Sub