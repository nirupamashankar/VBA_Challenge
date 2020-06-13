VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stockdata()


'setting variables

For Each WS In Worksheets
    
    Dim Ticker_name As String
    Dim Open_price As Double
    Dim Closing_price As Double
    Dim Price_Change As Double
    Dim Total_Volume As Double
    Dim Percent_change As Double
    Dim LastRow As Long
    Dim Summary_Table_Row As Long
    Dim i As Long
    
  'Setting initial values for variables that will change as loop runs
  
    Summary_Table_Row = 2
     
    Open_price = WS.Cells(2, 3).Value
    Total_Volume = 0
    
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'Setting headers for output table
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To LastRow
        
        If WS.Cells(i + 1, 1) <> WS.Cells(i, 1) Then
           
           Ticker_name = WS.Cells(i, 1).Value
           'print the ticker name to the appropriate column
           
           WS.Range("I" & Summary_Table_Row).Value = Ticker_name
           
           'Collect Closing price
           Closing_price = WS.Cells(i, 6).Value
           
           'calculate price change
            Price_Change = Closing_price - Open_price
            
            'print the price change in the appropriate column
             WS.Range("J" & Summary_Table_Row).Value = Price_Change
             
            
           'calculate percent change
             If Open_price = 0 Then
             Percent_change = 0
             
             ElseIf Price_Change <> 0 Then
                Percent_change = (Price_Change / Open_price)
             End If
             
            'populate in appropirate column
            WS.Range("K" & Summary_Table_Row).Value = Percent_change
            WS.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
             
             
            'calculate total volume
            Total_Volume = Total_Volume + WS.Cells(i, 7).Value
            'populate in appropriate column
            
            WS.Range("L" & Summary_Table_Row).Value = Total_Volume
            
            Price_Change = 0
            Percent_change = 0
            Total_Volume = 0
            Summary_Table_Row = Summary_Table_Row + 1
            Open_price = WS.Cells(i + 1, 3).Value
        
        Else
             'calculate total volume
             Total_Volume = Total_Volume + WS.Cells(i, 7).Value
        
        End If
        
        
    Next i
       
    'setting headers for second output table
    WS.Cells(1, 15).Value = "Ticker"
    WS.Cells(1, 16).Value = "Value"
    
    WS.Cells(2, 14).Value = "Greatest % Increase"
    WS.Cells(3, 14).Value = "Greatest % Decrease"
    WS.Cells(4, 14).Value = "Greatest Total Volume"
    
    
    
    Dim Max_change As Double
    Dim Min_change As Double
    Dim Max_Volume As Double
    lastrow_summary_table = WS.Cells(Rows.Count, 9).End(xlUp).Row
       
           
    'Conditional formatting color
    
    For j = 2 To lastrow_summary_table
            
                If WS.Cells(j, 10) < 0 Then
                WS.Cells(j, 10).Interior.ColorIndex = 3
                Else
                WS.Cells(j, 10).Interior.ColorIndex = 4
                End If
    Next j
    
                   
           
    For k = 2 To lastrow_summary_table
      If WS.Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & lastrow_summary_table)) Then
                WS.Cells(2, 15).Value = WS.Cells(k, 9).Value
                WS.Cells(2, 16).Value = WS.Cells(k, 11).Value
                WS.Cells(2, 16).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf WS.Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & lastrow_summary_table)) Then
                WS.Cells(3, 15).Value = WS.Cells(k, 9).Value
                WS.Cells(3, 16).Value = WS.Cells(k, 11).Value
                WS.Cells(3, 16).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf WS.Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & lastrow_summary_table)) Then
                WS.Cells(4, 15).Value = WS.Cells(k, 9).Value
                WS.Cells(4, 16).Value = WS.Cells(k, 12).Value
    End If
    Next k

Next WS

End Sub


