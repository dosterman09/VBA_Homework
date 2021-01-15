Attribute VB_Name = "Module1"
Sub TickerSymbol()

' Initial Variables for ALL worksheets
    Dim Ticker_symbol As String
    Dim yearly_change As Double
    Dim Percent_change As Double
    Dim Total_Volume As Double
    Dim max_percent As Double
    Dim min_percent As Double
    Dim lrow As Double
    Dim Range As Double
    
    
 ' Apply to all Worksheets
     Dim xIUp As Worksheet
     
    lrow = Cells(Rows.Count, "A").End(xIUp).Row
    Summary_table_row = 2
    Total_Volume = 0
    For i = 2 To lrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
    ' calculate yearly change
        yearlychange = Cells(i, 6).Value - Cells(2, 3).Value
        
    ' calculate percentage change
        percentchange = yearlychange / Cells(2, 3).Value
        
    ' set the ticker name
        Ticker_symbol = Cells(i, 1).Value
        
    ' add the volume total
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
    ' Print ticker symbol on row 2
        Range("i" & Summary_table_row).Value = Ticker_symbol
        
    ' Print total volume on row 2
        Range("j" & Summary_table_row).Value = Total_Volume
        
    ' Print yearly change value on row 2
        Range("k" & Summary_table_row).Value = yearlychange
    
    ' Print percent change value on row 2
        Range("L" & Summary_table_row).Value = percentchange
    
    ' If yearly change column value is positive format cell to be green
        If Range("K" & Summary_table_row).Value >= 0 Then
           Range("K" & Summary_table_row).Interior.ColorIndex = 40
           
    ' Or if change coloum value is negative format cell to be red
            ElseIf Range("K" & Summary_table_row).Value < 0 Then
                      Range("K" & Summary_table_row).Interior.ColorIndex = 30
                      
            End If
            
    ' add one to summary table row
            Summary_table_row = Summary_table_row + 1
            
    'Reset stock total volume
            Total_Volume = 0
            
            
    ' If Cell is same ticker symbol
             ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                Total_Volume = Total_Volume + Cells(i, 7).Value
                
            End If
            
    ' Range to determine greatest percentage_increase
            Set max_rng = Application.ActiveSheet.Range("L2:L999")
    ' Range to determine greatest percentage decrease
            Set min_rng = Application.ActiveSheet.Range("L2999")
    ' Range to determine greatest total volume
            Set max_volume_rng = Application.ActiveSheet.Range("J2;J999")
            
    
    ' Determine greatest value
            maxpercent = Application.WorksheetFunction.Max(max_rng)
    ' Determine minimum value
            minpercent = Application.WorksheetFunction.Min(max_rng)
    ' Determine greatest total volume
            maxvolume = Application.WorksheetFunction.Max(max_volume_rng)
            
    ' Row has greatest increase in percentage
      If Cells(i, 12).Value = maxpercent Then
      
    ' Print decrease
         Cells(2, 16).Value = minepercent
         Cells(3, 15).Value = Cells(i, 9).Value
         
            End If
            
      ' Row has greatest total volume
        If Cells(i, 10).Value = maxvolume Then
        
       ' Print greatest total volume to cell
        Cells(4, 16).Value = maxvolume
        Cells(4, 15).Value = Cells(i, 9).Value
        
        End If
        
        
        Next i
        
         
      
            
        
            
            
            
            
        
            
           
        
    
 

End Sub
