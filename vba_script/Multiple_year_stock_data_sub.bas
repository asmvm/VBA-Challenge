Attribute VB_Name = "Module1"

Sub VBAhomework()

' Apply script and formatting to all worksheets

Dim ws As Worksheet
Dim rowHeaders() As Variant

For Each ws In ThisWorkbook.Worksheets
ws.Activate

       ' Set column headers for summary table in each work sheet
        Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        Range("I1:L1").Select
        Selection.EntireColumn.AutoFit
        Range("I1:L1").Font.Bold = True
        
        Range("O1:P1").Value = Array("Ticker", "Value")
        Range("O1:P1").Select
        Selection.EntireColumn.AutoFit
        Range("O1:P1").Font.Bold = True
               
        rowHeaders() = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
        Range("N2:N4") = Application.Transpose(rowHeaders())
        Columns("N:P").Select
        Selection.EntireColumn.AutoFit
        Range("N2:N4").Font.Bold = True
               
  ' Set variables for Ticker summary table (modeled after Credit Card class exercise)
    Dim StockName As String
    Dim StockVolume As Double
            
    Dim OpeningPrice As Single
    Dim ClosingPrice As Single
    ' (YearlyChange returns an overflow error, edited for larger data types)
    Dim YearlyChange As Single
    Dim PercentChange As Double
    
    Dim lastrow As Long
        
    StockVolume = 0
    'YearlyChange = 0
    'PercentChange = 0
    'OpeningPrice = 0
    'ClosingPrice = 0
    
    Dim Ticker As Integer
    Ticker = 2
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
        For i = 2 To lastrow
        
        ' Test scripts: setting open price to resolve Divide by 0 error
            'OpeningPrice = Cells(Ticker, 3).Value
            'OpeningPrice = Range("C" & i + 1).Value
        
        'If ticker name not equal to previous ticker:
           
            If Cells(i + 1, 1).Value <> 0 And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                        
            StockName = Cells(i, 1).Value
            StockVolume = StockVolume + Cells(i, 7).Value
            OpeningPrice = Cells(Ticker, 3).Value
            ClosingPrice = Cells(i, 6).Value
        
        ' Set the calculation for Percent Change and YearlyChange.
        ' Percent Change is the Yearly Change/Opening Price
        ' Account for zeroes and jump to next value (can't divide by zero)
          
                If (OpeningPrice = 0 Or IsNull(OpeningPrice)) Then
                    OpeningPrice = Cells(i + 1, 3).Value
                End If
              
              YearlyChange = ClosingPrice - OpeningPrice
              PercentChange = YearlyChange / OpeningPrice
                          
            
        'Print Ticker Name, Yearly Change, Percent Change, Total Volume to summary table
                            
            Range("I" & Ticker).Value = StockName
            Range("L" & Ticker).Value = StockVolume
            
            Range("J" & Ticker).Value = YearlyChange
            Range("J" & Ticker).NumberFormat = "00.00"
            If Range("J" & Ticker).Value > 0 Then Range("J" & Ticker).Interior.ColorIndex = 4
            If Range("J" & Ticker).Value < 0 Then Range("J" & Ticker).Interior.ColorIndex = 3
                 
            Range("K" & Ticker).Value = PercentChange
            Range("K" & Ticker).NumberFormat = "00.00%"
           
            Ticker = Ticker + 1
        
        'Reset the expressions
            StockVolume = 0
            'Ticker = 0
          
                    
            Else
            
            'If ticker name is equal to each other:
             StockVolume = StockVolume + Cells(i, 7).Value
                         
            End If
            
        Next i
   
      'Print GreatestPercentIncrease, GreatestPercentDecrease, and Greatest Total Volume in SummaryTable
      
        'Dim lastinrange As Integer
        'lastinrange = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        Dim PctChg As Range
        Dim tsv As Range
        Dim Max As Double
        Dim Min As Double
        Dim tsvMax As Double
                  
            Set PctChg = ws.Range("K2:K10000") ' I hardcoded range b/c lastinrange or lastrow returned incorrect values
            Min = Application.WorksheetFunction.Min(PctChg)
            Max = Application.WorksheetFunction.Max(PctChg)
                
            Set tsv = ws.Range("L2:L10000") ' I hardcoded range b/c lastinrange or lastrow returned incorrect values
            tsvMax = Application.WorksheetFunction.Max(tsv)
                           
            Cells(2, 16).Value = Max
            Cells(3, 16).Value = Min
            Cells(4, 16).Value = tsvMax
           
           Range("P2:P3").NumberFormat = "0.00%"
            
           Dim rngPercent As Long
           Dim r As Integer
           rngPercent = ws.Cells(Rows.Count, 11).End(xlUp).Row
                                               
           
           For r = 2 To rngPercent
                If Max = ws.Cells(r, 11).Value Then
                ws.Cells(2, 15).Value = ws.Cells(r, 9).Value
                ElseIf Min = ws.Cells(r, 11).Value Then
                ws.Cells(3, 15).Value = ws.Cells(r, 9).Value
                End If
           Next r
                
           Dim rngTSV As Long
           Dim y As Integer
           rngTSV = ws.Cells(Rows.Count, 12).End(xlUp).Row
           Dim tsvHigh As Long
           
           For y = 2 To rngTSV
                If tsvMax = ws.Cells(y, 12).Value Then
                ws.Cells(4, 15).Value = ws.Cells(y, 9).Value
                End If
           Next y
                               
            
    Next ws

End Sub
    
