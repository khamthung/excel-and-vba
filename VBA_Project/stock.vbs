Sub stock_analyze()

  Dim Ticker_Symbol As String
  Dim Volumn_Total As Double
  Dim Summary_Table_Row As Integer
  Dim Opened_Price As Double
  Dim Closed_Price As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  Dim Greatest_Increase As Double
  Dim Greatest_Decrease As Double
  Dim Greatest_Total_Volumn As Double
  
  
For Each ws In Worksheets
    ws.Activate
    ' Set Cell Title
    Range("I" & "1").Value = "Ticker"
    Range("L" & "1").Value = "Total Stock Volumn"
    Range("J" & "1").Value = "Yearly Change"
    Range("K" & "1").Value = "Percent Change"
    Range("P" & "1").Value = "Ticker"
    Range("Q" & "1").Value = "Value"
    Range("O" & "2").Value = "Greatest % Increase"
    Range("O" & "3").Value = "Greatest % Decrease"
    Range("O" & "4").Value = "Greatest Total Volume"
    
    ' Initiate Value
    Volumn_Total = 0
    Previous_Yearly_Change = 0
    Yearly_Change = 0
    Opened_Price = 0
    Closed_Price = 0
    Summary_Table_Row = 2
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volumn = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_Symbol = Cells(i, 1).Value
            Closed_Price = Cells(i, 6).Value
            
            ' Calculation
            Yearly_Change = Closed_Price - Opened_Price
            Percent_Change = Yearly_Change / Opened_Price
            Volumn_Total = Volumn_Total + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            Range("L" & Summary_Table_Row).Value = Volumn_Total
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "Percent")
            
            
            ' Color Condition formatting
            If Yearly_Change >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            ' Searching for Greatest %
            If Percent_Change > Greatest_Increase Then
                Greatest_Increase = Percent_Change
                Range("Q" & "2").Value = Greatest_Increase
                Range("P" & "2").Value = Ticker_Symbol
            ElseIf Percent_Change < Greatest_Decrease Then
                Greatest_Decrease = Percent_Change
                Range("Q" & "3").Value = Greatest_Decrease
                Range("P" & "3").Value = Ticker_Symbol
            ElseIf Volumn_Total > Greatest_Total_Volumn Then
                Greatest_Total_Volumn = Volumn_Total
                Range("Q" & "4").Value = Greatest_Total_Volumn
                Range("P" & "4").Value = Ticker_Symbol
            End If
            
            ' Reset Value
            Summary_Table_Row = Summary_Table_Row + 1
            Closed_Price = 0
            Opened_Price = 0
            Volumn_Total = 0
            Percent_Change = 0
        Else
            Volumn_Total = Volumn_Total + Cells(i, 7).Value
            If Opened_Price = 0 Then
                Opened_Price = Cells(i, 3).Value
            End If
        End If
    Next i
    
    ws.Columns("A:Q").AutoFit
Next ws
End Sub




