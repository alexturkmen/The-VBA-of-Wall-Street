'This macro is for running the stockmarket macro in all the worksheets in this workbook
Sub loop_through_worksheets()

    Dim wbk As Workbook
    Dim ws As Worksheet
    
        Set wbk = ThisWorkbook
        Application.ScreenUpdating = False
    
            For Each ws In wbk.Worksheets
                ws.Select
                Call stockmarket
            Next ws
        Application.ScreenUpdating = True
        
End Sub

Sub stockmarket()

'Declaring Variables

Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim total As Double
Dim openprice As Double
Dim closeprice As Double
Dim row1 As Integer
Dim row2 As Integer
Dim row3 As Integer
Dim row4 As Integer
Dim row5 As Integer

' Storing Data

row1 = 2
row2 = 2
row3 = 2
total = 0

'Adding Titles

Range("I1").Value = "Ticker"
Range("J1").Value = "Closing Price"
Range("K1").Value = "Opening Price"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"

        'Formatting Title Cells
        
        Range("I1:N1").ColumnWidth = 18
        Range("I1:N1").Font.Bold = True
        Range("I1:N1").VerticalAlignment = xlCenter
        Range("I1:N1").HorizontalAlignment = xlCenter
        
            Range("P1:R1").ColumnWidth = 20
            Range("P1:R1").Font.Bold = True
            Range("P1:R1").VerticalAlignment = xlCenter
            Range("P1:R1").HorizontalAlignment = xlCenter
                
                'Formatting Other Cells
                
                    Range("I2:N290").VerticalAlignment = xlCenter
                    Range("I2:N290").HorizontalAlignment = xlCenter
                    
                        Range("P2:R4").VerticalAlignment = xlCenter
                        Range("P2:R4").HorizontalAlignment = xlCenter

'Calculating last row for Loop 1
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop 1 for opening price, closing price and total stock volume
For i = 2 To lastrow
    
    'Generating tickers, closing price, and total stock volume
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value
        Range("I" & row1).Value = ticker
               
            closeprice = Cells(i, 6).Value
            Range("J" & row1).Value = closeprice
            
                total = total + Cells(i, 7).Value
                Range("N" & row1).Value = total
                
                row1 = row1 + 1
                total = 0
                
                    'Generating opening price
                    ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
                    openprice = Cells(i, 3).Value
            
                    Range("K" & row2).Value = openprice
                    row2 = row2 + 1
                
                    Else
                    total = total + Cells(i, 7).Value
    End If
    
Next i

'Calculating last row for Loop 2
lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row

'Loop 2 for yearly change and percent change

For j = 2 To lastrow2

    'Generating Yearly Change
    yearlychange = Cells(j, 10).Value - Cells(j, 11).Value
    
    Range("L" & row3).Value = yearlychange
    row3 = row3 + 1
    
    'Generating percent change and making sure there is no error with the formula if the denominator is zero
    
        If Cells(j, 11).Value <> 0 Then
            percentchange = yearlychange / Cells(j, 11).Value
            
            Cells(j, 13).Value = percentchange
        
                ElseIf Cells(j, 11).Value = 0 Then
            
                Cells(j, 13).Value = 0
            
        End If

        'Conditional formatting for yearly change
            If Cells(j, 12).Value > 0 Then
                Cells(j, 12).Interior.ColorIndex = 4
                
                ElseIf Cells(j, 12).Value < 0 Then
                Cells(j, 12).Interior.ColorIndex = 3
            End If
     
                    'Conditional formatting for percent change
                    Cells(j, 13).NumberFormat = "0.00%"
                    
Next j

    'Deleting closing price and opening price columns
    Columns("J:K").EntireColumn.Delete
    
        'Challenge1 = generating greatest decrease/increase in percent change and greatest total stock volume
        Cells(2, 16) = Application.WorksheetFunction.Max(Range(Cells(1, 11), Cells(lastrow2, 11)))
        Cells(3, 16) = Application.WorksheetFunction.Min(Range(Cells(1, 11), Cells(lastrow2, 11)))
            Cells(2, 16).NumberFormat = "0.00%"
            Cells(3, 16).NumberFormat = "0.00%"
            
                Cells(4, 16) = Application.WorksheetFunction.Max(Range(Cells(1, 12), Cells(lastrow2, 12)))
                    
                    'Defining variables to lookup tickers for greatest values
                    Dim x As Double
                    Dim y As Double
                    Dim z As Double
                    
                        x = Application.WorksheetFunction.Max(Range(Cells(1, 11), Cells(lastrow2, 11)))
                        y = Application.WorksheetFunction.Min(Range(Cells(1, 11), Cells(lastrow2, 11)))
                        z = Application.WorksheetFunction.Max(Range(Cells(1, 12), Cells(lastrow2, 12)))
                            
                            'Loop 3 for looking up tickers and printing them on the appropriate cells
                            For k = 2 To lastrow2
                            
                                If Cells(k, 11).Value = x Then
                                Range("O2") = Cells(k, 9).Value
                                
                                        ElseIf Cells(k, 11).Value = y Then
                                        Range("O3") = Cells(k, 9).Value
                                        
                                            ElseIf Cells(k, 12).Value = z Then
                                            Range("O4") = Cells(k, 9).Value
                                            
                                End If
                                
                            Next k
                                    

End Sub