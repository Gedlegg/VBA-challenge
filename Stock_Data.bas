Attribute VB_Name = "Module1"
Sub Stock_Data()
    Dim ws As Worksheet
    Dim Last_Row As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Output_Row As Long
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Volume_Ticker As String
    Dim i As Long
    
    ' This code loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' This return the last row in column i
        Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Output_Row = 1
        
        ' This return column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through each row in the data
        For i = 2 To Last_Row
            ' This will check if the current ticker symbol is different from the previous one
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' This store the opening price
                Open_Price = ws.Cells(i, 3).Value
                ' This reset total volume
                Total_Volume = 0
            End If
            
            ' This store the closing price
            Close_Price = ws.Cells(i, 6).Value
            
            ' This will accumulate the total volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            ' Check if it's the last row of the current ticker or the last row of the worksheet
            If i = Last_Row Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' This will calculate yearly change
                Yearly_Change = Close_Price - Open_Price
                ' This will calculate percent change
                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                Else
                    Percent_Change = 0
                End If
                
                ' This will Round off percentage change to 2 decimal places
                Percent_Change = Round(Percent_Change, 2)
                
                ' This will return distinct Ticker Symbol and yearly change
                Output_Row = Output_Row + 1
                ws.Cells(Output_Row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Output_Row, 10).Value = Yearly_Change
                
                ' This will apply conditional color code based on yearly change
                If Yearly_Change > 0 Then
                    ws.Cells(Output_Row, 10).Interior.Color = RGB(0, 255, 0) ' Green color
                ElseIf Yearly_Change < 0 Then
                    ws.Cells(Output_Row, 10).Interior.Color = RGB(255, 0, 0) ' Red color
                End If
                
                ' This will return percentage change and toral volume
                ws.Cells(Output_Row, 11).Value = Percent_Change & "%" ' Concatenate a percent sign
                ws.Cells(Output_Row, 12).Value = Total_Volume
                
                ' This will return greatest increase, decrease, and volume
                If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Ticker = ws.Cells(i, 1).Value
                End If
                
                If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Ticker = ws.Cells(i, 1).Value
                End If
                
                If Total_Volume > Greatest_Volume Then
                    Greatest_Volume = Total_Volume
                    Greatest_Volume_Ticker = ws.Cells(i, 1).Value
                End If
            End If
        Next i
        
        ' This will return greatest increase,  greatest decrease, and greatest total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = Greatest_Increase_Ticker
        ws.Cells(3, 16).Value = Greatest_Decrease_Ticker
        ws.Cells(4, 16).Value = Greatest_Volume_Ticker
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 17).Value = Greatest_Increase & "%"
        ws.Cells(3, 17).Value = Greatest_Decrease & "%"
        ws.Cells(4, 17).Value = Greatest_Volume
    Next ws
End Sub


Sub Reset()
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Loop through each worksheet in the workbook
   For Each ws In ThisWorkbook.Sheets
        ' This will clear the contents of specific cells
        For i = 9 To 12
            ws.Cells(1, i).Value = "" ' Clear column headers
        
        Next i
        ' This will clear the contents and color formatting of the these output cells
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).Interior.Pattern = xlNone
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).Interior.TintAndShade = 0
        ws.Range("J2:J" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).Interior.PatternTintAndShade = 0
        ws.Range("I1:L" & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row).ClearContents
        ws.Cells(1, 15).Resize(4, 3).ClearContents ' Clear greatest increase, decrease, and volume cells
     Next ws
    
End Sub


