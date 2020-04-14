Sub StockMarketAnalysis():
    ' Loop Through Worksheets
    For Each ws In Worksheets

        Dim Ticker_Name As String
        Dim Last_Row As Long
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double
        Dim Yearly_Change As Double
        Dim Previous_Amount As Long
        Previous_Amount = 2
        Dim Percent_Change As Double
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        Dim Last_Row_Value As Long
        Dim Greatest_Total_Volume As Double
        Greatest_Total_Volume = 0

        'Get Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To Last_Row

            ' Add To Ticker Total Volume
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            ' Check If Still Same Ticker Name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get Ticker Name
                Ticker_Name = ws.Cells(i, 1).Value
                ' Show The Ticker Name In The Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ' Show The Ticker Total Amount To The Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                ' Reset Total
                Total_Ticker_Volume = 0

                ' Get Yearly Open, Yearly Close and Yearly Change Name
                Yearly_Open = ws.Range("C" & Previous_Amount)
                Yearly_Close = ws.Range("F" & i)
                Yearly_Change = Yearly_Close - Yearly_Open
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                ' Percent Change
                If Yearly_Open = 0 Then
                    Percent_Change = 0
                Else
                    Yearly_Open = ws.Range("C" & Previous_Amount)
                    Percent_Change = Yearly_Change / Yearly_Open
                End If
                ws.Range("K" & Summary_Table_Row).Number_Format = "0.00%"
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change

                ' Make Green or Red
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color_Index = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color_Index = 3
                End If
            
                ' Add another to the Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1
                Previous_Amount = i + 1
                End If
            Next i

    Next ws

End Sub
