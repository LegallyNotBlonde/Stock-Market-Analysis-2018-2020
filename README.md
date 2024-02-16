# Stock-Market-Analysis for 2018-2020#

##Analyzing muti-stock price and volume change for the years of 2018. 2019, and 2020##

* Conducting an in-depth analysis of multi-stock price dynamics and volume fluctuations throughout the fiscal years of 2018, 2019, and 2020 constitutes the primary objective of this project. The analytical framework involves parsing through extensive datasets encompassing stock market metrics over the specified period.
* The implemented script iterates through rows of data to discern patterns and unveil insights regarding the performance of individual stocks.
* Key calculated metrics include yearly price change, both in total value and percentage terms. Furthermore, the script identifies and accentuates noteworthy trends such as the stock with the highest percentage increase, signifying optimal performance, as well as pinpointing the most significant decliner in percentage terms. Additionally, the analysis discerns the stock boasting the highest volume, indicative of either superior market liquidity or enduring market sustainability.
* Links to screenshots and VBA script
  
1. https://github.com/LegallyNotBlonde/Stock-Market-Analysis-2018-2020/blob/main/Screen%20shot%20results%20%20for%202018.png
2. https://github.com/LegallyNotBlonde/Stock-Market-Analysis-2018-2020/blob/main/Screen%20shot%20results%20%20for%202019.png
3. https://github.com/LegallyNotBlonde/Stock-Market-Analysis-2018-2020/blob/main/Screen%20shot%20results%20%20for%202020.png
4.  VBA script link to the file: https://github.com/LegallyNotBlonde/Stock-Market-Analysis-2018-2020/blob/main/Stock%20Change%20VBA%20Script%20file%20-%202%20subrutines.bas



* Script:

Sub StockChange():
For Each ws In Worksheets

'Set variables

Dim Ticker, Total_Stock_Volume, Max_Volume, MaxIncrTicker, MaxDecrTicker, MaxVolTicker As String

Dim Open_Price, Close_Price, Yearly_Change, Percentage_Change, Gr_Incr, Gr_Decr As Double

Dim i, x, y As Integer

'set starting values for calculated variables and indexes
'second index y helped to have values entered at the top line by line and not on the lines where ticket # was changed

Yearly_Change = 0

Percentage_Change = 0

Total_Stock_Volume = 0

    x = 2
    y = 2

'Set percentage change in correct format

ws.Range("K2:K5000").NumberFormat = "0.00%"

ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("Q4").NumberFormat = "0"


'Creating a loop to go through all cells with values


For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

Open_Price = ws.Cells(x, 3).Value

Total_Stock_Volume = ws.Cells(i, 7).Value + Total_Stock_Volume


If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    'skip to the next row
    

Else:
    ws.Cells(y, 9).Value = ws.Cells(i, 1).Value
'Entering symbols for ticker

'entering value for Yearly_Change, percentage change, and Total stock volume

        Yearly_Change = ws.Cells(i, 6).Value - Open_Price
        
        ws.Cells(y, 10).Value = Yearly_Change
        
        Percentage_Change = (Yearly_Change) / Open_Price
        
        ws.Cells(y, 11).Value = Percentage_Change
        
        ws.Cells(y, 12).Value = Total_Stock_Volume
        
        
'fund max stock value, max % increase and max % decrease

        Max_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        
        ws.Cells(4, 17) = Max_Volume
        
        Gr_Incr = Application.WorksheetFunction.Max(ws.Range("K:K"))

        ws.Cells(2, 17) = Gr_Incr
      
        Gr_Decr = Application.WorksheetFunction.Min(ws.Range("K:K"))
       
        ws.Cells(3, 17) = Gr_Decr
        
'use index and match function to enter which stock values have the highest increase, decrease, and volume
        
        ws.Range("P2") = MaxIncrTicker
        
        ws.Range("P3") = MaxDecrTicker
       
        ws.Range("P4") = MaxVolTicker
        
        
        MaxIncrTicker = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Gr_Incr, ws.Range("K:K"), 0))
       
        MaxDecrTicker = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Gr_Decr, ws.Range("K:K"), 0))
        
        MaxVolTicker = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Max_Volume, ws.Range("L:L"), 0))
        
        
        x = i + 1
        y = y + 1
        Yearly_Change = 0
        Percentage_Change = 0
        Total_Stock_Volume = 0


    End If

Next i
'insert text in headers

ws.Range("I1").Value = "Ticker"

ws.Range("J1").Value = "Yearly Change"

ws.Range("K1").Value = "Percent Change"

ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"

ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"

ws.Range("O3").Value = "Greatest % Decrease"

ws.Range("O4").Value = "Greatest Total Volume"


Next ws


End Sub


Sub FormattingColors()


Dim ws As Worksheet
    
Dim column As Integer
    
Dim Percentage_Change As Double

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
    
        column = 11 ' Assuming I want to apply color to column K (11th column) on every sheet

        ' Loop through rows in the current worksheet
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
            ' Set colors based on values (green for >0, red for <0, and default (no color) for changes equal to 0)
            'the task did not specify, which color should the cells have if the change is equal to 0
            
            If ws.Cells(i, column).Value > 0 Then
                ws.Cells(i, column).Interior.ColorIndex = 4 ' Green
                
            ElseIf ws.Cells(i, column).Value < 0 Then
                ws.Cells(i, column).Interior.ColorIndex = 3 ' Red
                
            Else
                ws.Cells(i, column).Interior.ColorIndex = 2 ' Default color
            End If
        Next i
    Next ws
End Sub






