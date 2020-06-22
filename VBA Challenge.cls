VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBAStocks():

Dim Ticker As String
Dim Number_Tickers As Integer
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Last_Row As Long
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Increase_Ticker As String
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Percent_Decrease_Ticker As String
Dim Greatest_Total_Volume As Double
Dim Greatest_Total_Volume_Ticker As String

'Loop through all the worksheets in the workbook
For Each ws In Worksheets

        'Make the worksheet active
            ws.Activate
    
        'Locate the last row in each worksheet
            LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Set the values for each worksheet
            Number_Tickers = 0
            Ticker = ""
            Opening_Price = 0
            Percent_Change = 0
            Total_Stock_Volume = 0
            Yearly_Change = 0
            
        'Omitting first row that contains header, loop through the rest of the tickers
        For i = 2 To LastRow
        
            'Determine Ticker Symbol
            Ticker = Cells(i, 1).Value
            
            'Determine Opening Price
            If Opening_Price = 0 Then
                Opening_Price = Cells(i, 3).Value
                
            End If
            
            'Determine Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            'Ticker different than listed
            If Cells(i + 1, 1).Value <> Ticker Then
                'Increase Number_Tickers if not listed
                Number_Tickers = Number_Tickers + 1
                Cells(Number_Tickers + 1, 9) = Ticker
            
                'Determine Closing Price
                Closing_Price = ws.Cells(i, 6).Value
                
                'Determine Yearly Change
                Yearly_Change = Closing_Price - Opening_Price / Opening_Price
                
                'Calculate Yearly Change
                Cells(Number_Tickers + 1, 10).Value = Yearly_Change
                
                'If Yearly Change Greater than 0, Green Shade
                If Yearly_Change > 0 Then
                    Cells(Number_Tickers + 1, 11).Interior.ColorIndex = 4
                'If Yearly Change Less than 0, Red Shade
                ElseIf Yearly_Change < 0 Then
                    Cells(Number_Tickers + 1, 11).Interior.ColorIndex = 3
                End If
                
                'Determining the value of Percent Change for Ticker
                If Opening_Price = 0 Then
                Percent_Change = 0
                Else
                    Percent_Change = (Yearly_Change / Opening_Price)
                End If
                
                'Determing Percent_Change as Percent
                Cells(Number_Tickers + 1, 11).Value = Format(Percent_Change, "Percent")
    
                'Set variable to revert Opening Price to 0 when different ticker in list
                Opening_Price = 0
                
                'Set variable to include Total Stock Value in each worksheet
                Cells(Number_Tickers + 1, 12).Value = Total_Stock_Volume
                
                'Set variable to revert Total Stock Volume to 0 when different ticker in list
                Total_Stock_Volume = 0
                End If
                
            Next i
            
            'Set variable to include table that displays Greatest Percent Increase, Decrease, Total Volume, Ticker & Value
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            
            'Locate the last row in each worksheet
            LastRow = Cells(Rows.Count, "I").End(xlUp).Row
            
            'Set variables to determine list in worksheet
            Greatest_Percent_Increase = Cells(2, 11).Value
            Greatest_Percent_Increase_Ticker = Cells(2, 9).Value
            Greatest_Percent_Decrease = Cells(2, 11).Value
            Greatest_Percent_Decrease_Ticker = Cells(2, 9).Value
            Greatest_Stock_Volume = Cells(2, 12).Value
            Greatest_Stock_Volume_Ticker = Cells(2, 9).Value
            
            'Omitting first row that contains header, loop through the rest of the tickers
            For i = 2 To LastRow
            
                'Determine Ticker that contains Greatest Percent Increase
                If Cells(i, 11).Value > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = Cells(i, 11).Value
                Greatest_Percent_Increase_Ticker = Cells(i, 9).Value
                End If
                
                'Determine Ticker that contains Greatest Percent Decrease
                If Cells(i, 11).Value < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = Cells(i, 11).Value
                Greatest_Percent_Decrease_Ticker = Cells(i, 9).Value
                End If
                
                'Determine Ticker that contains Greatest Stock Volume
                If Cells(i, 12).Value > Greatest_Stock_Volume Then
                Greatest_Stock_Volume = Cells(i, 12).Value
                Greatest_Stock_Volume_Ticker = Cells(i, 9).Value
                End If
            
            Next i
            
          
    Next ws
    
End Sub


