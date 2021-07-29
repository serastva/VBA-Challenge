Sub Multi_year_stock_final()


For Each ws In Worksheets
    ws.Activate



'This will create the column names and formats column width
'------------------------------------------------------------------------
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greater % Increase"
Range("O3").Value = "Greater % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Columns("I:Q").AutoFit

'Select the last cell with values
'------------------------------------------------------------------------

Dim Last_Row As Long
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

'------------------------------------------------------------------------
'Sets Rows and Columns

Dim Row As Integer
Row = 2

Dim Column As Integer
Column = 1


'Defining Initial Conditions
Dim Total_Volume As Double

Total_Volume = 0

Dim PriceOpen As Double
PriceOpen = Cells(2, Column + 2).Value


Greatest_Percent_Increase = 0
Greatest_Percent_Decrease = 0
Greatest_Total_Volume = 0


For I = 2 To Last_Row

Total_Volume = Total_Volume + Cells(I, Column + 6).Value
Cells(Row, "L").Value = Total_Volume


'------------------------------------------------------------------------
'Ticker Names
    If Cells(I + 1, Column).Value <> Cells(I, Column).Value Then
    
        'Set Ticker Name
        TickerName = Cells(I, Column).Value
        
        'Print Ticker Name in the Table
        Cells(Row, "I").Value = TickerName
'------------------------------------------------------------------------
'Yearly Change

        'Set Yearly Name
        Dim PriceClose
        PriceClose = Cells(I, Column + 5).Value
        
        'Print Yearly Change in the Table
        Dim YearlyChange As Double
        YearlyChange = PriceClose - PriceOpen
        Cells(Row, "J") = YearlyChange
'-------------------------------------------------------------------------
'Percent Change
        
        'Set Percent Change
        
        If PriceOpen <> 0 Then
        
            Dim PercentChange As Double
            PercentChange = (PriceClose - PriceOpen) / PriceOpen
        
            'Print Percent Change
            Cells(Row, "K") = Format(PercentChange, "percent")
        
            
            
        Else
           Cells(2, 11).Value = Format(0, "Percent")
          
          
        End If
          
        
'--------------------------------------------------------------------------
'Sets colors to the Yearly Change tab
         If YearlyChange > 0 Then
            Cells(Row, "J").Interior.Color = RGB(0, 255, 0)
         Else
            Cells(Row, "J").Interior.Color = RGB(255, 0, 0)
       
         End If
        
'--------------------------------------------------------------------------
            If PercentChange > 0 And PercentChange > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = Cells(Row, "K").Value
                Greatest_Percent_Change = PercentChange
                Cells(2, "Q").Value = Greatest_Percent_Change
                Cells(2, "Q").NumberFormat = "0.00%"
                Cells(2, "P").Value = Cells(Row, "I")
            End If
            
            
            If PercentChange < 0 And PercentChange < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = Cells(Row, "K").Value
                Greatest_Percent_Change = PercentChange
                Cells(3, "Q").Value = Greatest_Percent_Change
                Cells(3, "Q").NumberFormat = "0.00%"
                Cells(3, "P").Value = Cells(Row, "I")
            End If
            
            'Gathering the greatest total volume
            If Total_Volume > 0 And Total_Volume > Greatest_Total_Volume Then
                Greatest_Total_Volume = Cells(Row, "L").Value
                Greatest_Total_Change = Total_Volume
                Cells(4, "Q").Value = Greatest_Total_Change
                Cells(4, "P").Value = Cells(Row, "I")
            End If








        
         'Add 1 to Row
        Row = Row + 1
        
        'Reset
        PriceOpen = Cells(I + 1, Column + 2).Value
        Total_Volume = 0
        
        
         
        End If
    
    Next I
    
Next ws
    
End Sub

