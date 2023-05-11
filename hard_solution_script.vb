Sub tickerSummary()
Dim ws As Integer
Dim sheet_count As Integer
sheet_count = ActiveWorkbook.Worksheets.Count
For ws = 1 To sheet_count
    ActiveWorkbook.Worksheets(ws).Activate
    Dim num_row As Long
    num_row = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Ticker_name As String
    Dim Open_date As Double
    Dim Open_price As Double
    Dim Close_date As Double
    Dim Close_price As Double
    Dim Yearly_change As Double
    Dim great_increase As Double
    great_increase = 0
    Dim great_decrease As Double
    great_decrease = 0
    Dim great_total_vol As Double
    great_total_vol = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim i As Long
    Dim j As Long

    [B:B].Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With

    Range(Cells(1, 1), Cells(num_row, 7)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
    Range("I1").Value = "Ticker name"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
  
        For i = 2 To num_row
        
        If Application.WorksheetFunction.CountIf(Range(Cells(1, 9), Cells(Summary_Table_Row, 9)), Cells(i, 1).Value) = 0 Then
            
            Ticker_name = Cells(i, 1).Value
           
            Cells(Summary_Table_Row, 9).Value = Ticker_name
            
            Cells(Summary_Table_Row, 12).Value = Application.WorksheetFunction.SumIf(Range("A:A"), Ticker_name, Range("G:G"))
           
            Open_date = Cells(i, 2).Value
            Close_date = Cells(i, 2).Value
            Open_price = Cells(i, 3).Value
            Close_price = Cells(i, 6).Value
            Summary_Table_Row = Summary_Table_Row + 1
       
        ElseIf Cells(i, 2).Value < Open_date Then
            Open_price = Cells(i, 3).Value
    
        ElseIf Cells(i, 2).Value > Close_date Then
            Close_price = Cells(i, 6).Value
        End If
    
        Yearly_change = Close_price - Open_price
        
        Cells(Summary_Table_Row - 1, 10).Value = Yearly_change
       
        If Yearly_change > 0 Then
            Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 4
        ElseIf Yearly_change < 0 Then
            Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 3
        End If
        
        Cells(Summary_Table_Row - 1, 11).Value = Close_price / Open_price - 1
      
        Cells(Summary_Table_Row - 1, 11).NumberFormat = "0.00%"
      Next i
  
    Dim num_row_new As Long
    num_row_new = Cells(Rows.Count, 9).End(xlUp).Row

        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
      
        For j = 2 To num_row_new
        
        If Cells(j, 11).Value > great_increase Then
            great_increase = Cells(j, 11).Value
            Range("Q2").Value = great_increase
            Range("Q2").NumberFormat = "0.00%"
            Range("P2").Value = Cells(j, 9).Value
        End If
    
        If Cells(j, 11).Value < great_decrease Then
            great_decrease = Cells(j, 11).Value
            
            Range("Q3").Value = great_decrease
            Range("Q3").NumberFormat = "0.00%"
            Range("P3").Value = Cells(j, 9).Value
        End If

        If Cells(j, 12).Value > great_total_vol Then
            great_total_vol = Cells(j, 12).Value
            Range("Q4").Value = great_total_vol
            Range("P4").Value = Cells(j, 9).Value
        End If
        Next j
   
    Columns("I:Q").AutoFit
    num_row = 0
Next ws
End Sub