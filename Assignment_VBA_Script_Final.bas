Attribute VB_Name = "Module11"

Public Sub SRonallWorksheets()
                'This Sub commences the StockReport on all WorkSheets - It starts here... :)

Dim WS_Count As Integer
Dim I As Integer

WS_Count = ActiveWorkbook.Worksheets.Count

For I = 1 To WS_Count
    Worksheets(I).Select
    Call StockReport
    Range("I2").Select
        
Next I
    Worksheets(1).Select
    
End Sub

Public Sub StockReport()

                'This is the main sub for running the Stock Report on the Worksheet

Call ReportsTable
Call TickerIdentifier
Call Yearly_Percent_Change
Call YearlyChangeCheck
Call Total_Stock_Vol
Call GreatestIncrease
Call GreatestDecrease
Call GreatestVolume


End Sub

Public Sub ReportsTable()

' This creates and fits the headings for the main report

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("I1:L1").EntireColumn.AutoFit

'This creates and fits the table for the bonus report

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("O1:P4").EntireColumn.AutoFit
    
    
End Sub

Public Sub YearlyChangeCheck()

            'This Sub is to calulate the yearly price change for each Ticker and then applyies cell formatting
            'with green for postive, red for negative and cyan for nil

Dim DataCount, RowNum As Double

DataCount = Cells(Rows.Count, 11).End(xlUp).Row         'Counting all the data in the columns

For RowNum = 1 To DataCount - 1

    If Range("J" & RowNum + 1) > 0 Then
        Range("J" & RowNum + 1).Interior.Color = vbGreen
    ElseIf Range("J" & RowNum + 1) = 0 Then
        Range("J" & RowNum + 1).Interior.Color = vbCyan
    Else
        Range("J" & RowNum + 1).Interior.Color = vbRed
        
    End If

Next RowNum

End Sub

Public Sub TickerIdentifier()

            'This sub identifies the Tickers by looping and checking each row against the
            'previous line and when they dont match it prints the ticker in the Ticker column

Dim DataCount, RowNum, TickerCount As Double

TickerCount = 1

DataCount = Cells(Rows.Count, 1).End(xlUp).Row
        
For RowNum = 1 To DataCount - 1
        
    If Range("A" & RowNum).Value <> Range("A" & RowNum + 1) Then
               Range("I" & TickerCount + 1).Value = Range("A" & RowNum + 2).Value
               TickerCount = TickerCount + 1
    End If

Next RowNum

End Sub

Public Sub Yearly_Percent_Change()

            'This sub calculates the difference between the open value at 2/1/xxx
            'and closing value at 31/12/xx
            'i.e. Closing Value at 2/1/2020 less Opening value at 31/12/2020 = Yearly Change

            'This sub will also calcualte the percentage between the open
            'and closing values
            'i.e The percentage difference = difference caluclated above divided
            'into the Opening Value

Dim TickerRepCount, TickRepRowNum, DataCount, DataRowNum, OpenV, CloseV, Difference As Double
Dim FirstDate, LastDate As String

DataCount = Cells(Rows.Count, 1).End(xlUp).Row
TickerRepCount = Cells(Rows.Count, 9).End(xlUp).Row

FirstDate = Range("B2").Value
LastDate = Cells(Rows.Count, 2).End(xlUp).Value

For TickerRepRowNum = 1 To TickerRepCount - 1
                 
    For DataRowNum = 1 To DataCount - 1
        
        If Range("I" & TickerRepRowNum + 1).Value = Range("A" & DataRowNum + 1).Value Then
            If Range("B" & DataRowNum + 1).Value = FirstDate Then
                OpenV = Range("C" & DataRowNum + 1).Value
            ElseIf Range("B" & DataRowNum + 1).Value = LastDate Then
                CloseV = Range("F" & DataRowNum + 1).Value
            End If
        End If
        
    Next DataRowNum
    
    Difference = CloseV - OpenV
    Range("J" & TickerRepRowNum + 1).Value = Difference
    Range("K" & TickerRepRowNum + 1).Value = Difference / OpenV
    Range("K" & TickerRepRowNum + 1).Select
    Selection.NumberFormat = "0.00%"
      
Next TickerRepRowNum

End Sub

Public Sub Total_Stock_Vol()

                'This sub calculates the total Stock volume traded for the year for each Ticker through a loop

Dim TickerRepCount, TickRepRowNum, DataCount, DataRowNum, StockVol As Double

DataCount = Cells(Rows.Count, 1).End(xlUp).Row
TickerRepCount = Cells(Rows.Count, 9).End(xlUp).Row
StockVol = 0

For TickerRepRowNum = 1 To TickerRepCount - 1
    
        StockVol = O
        
 For DataRowNum = 1 To DataCount - 1

    If Range("I" & TickerRepRowNum + 1).Value = Range("A" & DataRowNum + 1).Value Then
        StockVol = StockVol + Range("G" & DataRowNum + 1)
    End If
    
  Next DataRowNum

    Range("L" & TickerRepRowNum + 1).Value = StockVol
    Range("L" & TickerRepRowNum + 1).Select
    Selection.NumberFormat = "0"
    
Next TickerRepRowNum

End Sub

Public Sub GreatestIncrease()

                'This Sub analyses the Percent Change column and identifies the Ticker with biggest
                'percentage increase over the year

Dim DataCount, DataRowNum, IncreaseV As Double
Dim Ticker As String

IncreaseV = Range("K2").Value
Ticker = Range("I2").Value

DataCount = Cells(Rows.Count, 11).End(xlUp).Row

For DataRowNum = 1 To DataCount - 1

    If Range("K" & DataRowNum + 2).Value > IncreaseV Then
        IncreaseV = Range("K" & DataRowNum + 2).Value
        Ticker = Range("I" & DataRowNum + 2).Value
    End If
    
Next DataRowNum

    Range("P2").Value = Ticker
    Range("Q2").Value = IncreaseV
    Range("Q2").Select
    Selection.NumberFormat = "0.00%"

End Sub

Public Sub GreatestDecrease()

                'This Sub anaylses the Percent Change column and identifies the Ticker with biggest
                'percentage decrease over the year

Dim DataCount, DataRowNum, DecreaseV As Double
Dim Ticker As String

DecreaseV = Range("K2").Value
Ticker = Range("I2").Value

DataCount = Cells(Rows.Count, 11).End(xlUp).Row

For DataRowNum = 1 To DataCount - 1

    If Range("K" & DataRowNum + 2).Value < DecreaseV Then
        DecreaseV = Range("K" & DataRowNum + 2).Value
        Ticker = Range("I" & DataRowNum + 2).Value
    End If
    
Next DataRowNum

    Range("P3").Value = Ticker
    Range("Q3").Value = DecreaseV
    Range("Q3").Select
    Selection.NumberFormat = "0.00%"

End Sub

Public Sub GreatestVolume()

                'This Sub analyses the Volume column and identifies the Ticker with biggest
                'total volume stock volume over the year

Dim DataCount, DataRowNum, StockV As Double
Dim Ticker As String

StockV = Range("L2").Value
Ticker = Range("I2").Value

DataCount = Cells(Rows.Count, 11).End(xlUp).Row

For DataRowNum = 1 To DataCount - 1

    If Range("L" & DataRowNum + 2).Value > StockV Then
        StockV = Range("L" & DataRowNum + 2).Value
        Ticker = Range("I" & DataRowNum + 2).Value
    End If
    
Next DataRowNum

    Range("P4").Value = Ticker
    Range("Q4").Value = StockV
    Range("Q4").Select
    Selection.NumberFormat = "0"
    Range("Q1:Q4").EntireColumn.AutoFit

End Sub
