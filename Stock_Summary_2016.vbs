Sub StockSummary2016()

 ' Set a variable for specifying the column of interest
 
Cells(1, 9).Value = "Ticker"

Cells(1, 10).Value = "YOY Price Change"

Cells(1, 11).Value = "YOY Percent Change"

Cells(1, 12).Value = "Total Stock Volume"

Cells(2, 15).Value = "Greatest % Increase"
        
Cells(3, 15).Value = "Greatest % Decrease"
        
Cells(4, 15).Value = "Greatest Total Volume"

Cells(1, 16).Value = "Ticker"
        
Cells(1, 17).Value = "Value"

Dim Ticker As String

Dim Volume As Double

Volume = 0

Dim Open_Price As Double

Dim Close_Price As Double

Dim Price_Change As Double

Dim Price_Percent_Change As Double

Dim Greatest_Increase As Double

Dim Greatest_Decrease As Double

Dim Greatest_Vol As Double

Dim Increse_Ticker As String

Dim Decrease_Ticker As String

Dim Vol_Ticker As String

Dim wks As Worksheet

Set wks = ActiveSheet

Dim RowRange As Range

Dim LastRow As Long

LastRow = wks.Cells(wks.Rows.Count, 1).End(xlUp).Row

Dim Summary As Integer

Summary = 2

  For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        
        Close_Price = Cells(i, 6).Value
        
        Open_Price = Cells(i - 261, 3).Value
        
        Price_Change = Close_Price - Open_Price
                
        Price_Percent_Change = (Close_Price - Open_Price) / Close_Price
        
        Volume = Volume + Cells(i, 7).Value
        
        Range("I" & Summary).Value = Ticker
        
        Range("J" & Summary).Value = Price_Change
            
            Range("J1").EntireColumn.AutoFit
            
            If Price_Change < 0 Then
            
            Range("J" & Summary).Interior.ColorIndex = 3
            
            ElseIf Price_Change > 0 Then
            
            Range("J" & Summary).Interior.ColorIndex = 4
            
            End If
        
        Range("K" & Summary).Value = Price_Percent_Change
        
            Range("K" & Summary).NumberFormat = "0.00%"
            
            Range("K1").EntireColumn.AutoFit
        
        Range("L" & Summary).Value = Volume
        
            Range("L1").EntireColumn.AutoFit
        
            Range("O1").EntireColumn.AutoFit
        
        Summary = Summary + 1
        
        Volume = 0
        
        Else
        
        Volume = Volume + Cells(i, 7).Value
         
    End If
  
  Next i
  
  Greatest_Increase = Application.WorksheetFunction.Max(Range("K:K"))
  
    Cells(2, 17).Value = Greatest_Increase
    
    Cells(2, 17).NumberFormat = "0.00%"
  
  Greatest_Decrease = Application.WorksheetFunction.Min(Range("K:K"))
  
    Cells(3, 17).Value = Greatest_Decrease
    
    Cells(3, 17).NumberFormat = "0.00%"
  
  Greatest_Vol = Application.WorksheetFunction.Max(Range("L:L"))
  
    Cells(4, 17) = Greatest_Vol
  
  ' Bonus ticker for the greatest increase (%) in stock price
For i = 2 To LastRow

    If Cells(i, 11) = Greatest_Increase Then
    
    Cells(2, 16).Value = Cells(i, 9).Value
    
    End If
    
Next i

' Bonus ticker for the greatest decrease (%) in stock price
For i = 2 To LastRow

    If Cells(i, 11) = Greatest_Decrease Then
    
    Cells(3, 16).Value = Cells(i, 9).Value
    
    End If
    
Next i

' Bonus ticker for the greatest volume ($)
For i = 2 To LastRow

    If Cells(i, 12) = Greatest_Vol Then
    
    Cells(4, 16).Value = Cells(i, 9).Value
    
    End If
    
Next i
    
End Sub