Attribute VB_Name = "Module1"
Sub Ticker()
    
    Dim NoRows As Long
    Dim ws As Worksheet
    Dim Ticker_Name As String
    Dim Ticker_Volume As Double
    Dim Open_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Summary_Table_Row As Integer
    
    NoRows = Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_Volume = 0
    Open_Price = ActiveWorkbook.Sheets(1).Cells(2, 3).Value
    Summary_Table_Row = 2

    
    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker Name"
        ws.Range("L1").Value = "Total Volume"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Columns(11).NumberFormat = "0.00%"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P1").Value = "Ticker Name"
        ws.Range("Q1").Value = "Value"
        ws.Range("O4").Value = "Most Volume"
        
        Ticker_Volume = 0
        Open_Price = Cells(2, 3).Value
        Summary_Table_Row = 2
        
        For i = 2 To NoRows
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker_Name = ws.Cells(i, 1).Value
                Closing_Price = ws.Cells(i, 6).Value
                
                If i = 2 Then
                    Closing_Price = ws.Cells(2, 6).Value
                End If
                
                Yearly_Change = Closing_Price - Open_Price
                Percent_Change = Yearly_Change / Open_Price
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                ws.Range("i" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                If Yearly_Change >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
        
                
                Open_Price = ws.Cells(i + 1, 3).Value
                Summary_Table_Row = Summary_Table_Row + 1
                Ticker_Volume = 0

            Else
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
            End If
            
        Next i
               
    Dim Max_Increase_Ticker As String
    Dim Max_Decrease_Ticker As String
    Dim Max_Volume_Ticker As String

    Dim Max_Increase As Double
    Dim Max_Decrease As Double
    Dim Max_Volume As Double

    Max_Increase_Ticker = ""
    Max_Decrease_Ticker = ""
    Max_Volume_Ticker = ""
    
    Max_Increase = 0
    Max_Decrease = 0
    Max_Volume = 0

    For j = 2 To Range("K1").End(xlDown).Row
        If ws.Cells(j, 11).Value > Max_Increase Then
            Max_Increase = ws.Cells(j, 11).Value
            Max_Increase_Ticker = ws.Cells(j, 9).Value
    End If
    
    If ws.Cells(j, 11).Value < Max_Decrease Then
        Max_Decrease = ws.Cells(j, 11).Value
        Max_Decrease_Ticker = ws.Cells(j, 9).Value
    End If
    
    If ws.Cells(j, 12).Value > Max_Volume Then
        Max_Volume = ws.Cells(j, 12).Value
        Max_Volume_Ticker = ws.Cells(j, 9).Value
    End If
    
        ws.Range("P2").Value = Max_Increase_Ticker
        ws.Range("P3").Value = Max_Decrease_Ticker
        ws.Range("P4").Value = Max_Volume_Ticker

        ws.Range("Q2").Value = Max_Increase
        ws.Range("Q3").Value = Max_Decrease
        ws.Range("Q4").Value = Max_Volume
        ws.Range("Q4").NumberFormat = "0"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ws.Range("I2:Q2").EntireColumn.Autofit
        
        Next j
        
    Next ws
        
End Sub

Sub Autofit()
   
   Dim ws As Worksheet
   
   For Each ws In Worksheets
   
   ws.Range("I2:Q2").EntireColumn.Autofit
   
   Next ws
    
End Sub

