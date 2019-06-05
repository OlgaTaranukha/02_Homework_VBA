Option Explicit

Type TypeOfResult
   ticker       As String
   price_begin  As Currency ' Single
   price_end    As Currency ' Single
   volume       As Double
End Type


Sub Calculate_Results()
    Dim current_worksheet As Worksheet
    
    ' Do for each worksheet in this workbook
    For Each current_worksheet In ThisWorkbook.Worksheets
        Call Calculate_Subtotals(current_worksheet)
    Next current_worksheet

    MsgBox ("Work is done")
End Sub

Sub Calculate_Subtotals(ws As Worksheet)

    Dim arr_results() As TypeOfResult
    Dim l_Row, l_RowLast, l_ColLast As Long
    Dim s_ColLastLetter As String
    Dim v_max, v_min, v_max_vol As Variant
    Dim d_price_change As Double
    Dim i, j As Long
    Dim r As range
    
    ' Do for current worksheet
    With ws
         .Select
         
         ' Header for total table
         .range("I1").Value = "Ticker"
         .range("J1").Value = "Yearly Change"
         .range("K1").Value = "Percent Change"
         .range("L1").Value = "Total Stock Volume"
         
         ' Header for Max % Increase , Max, Max % Decrease and Max Volume
         .range("O2").Value = "Greatest % Increase"
         .range("O3").Value = "Greatest % Decrease"
         .range("O4").Value = "Greatest Total Volume"
         .range("P1").Value = "Ticker"
         .range("Q1").Value = "Value"
         
         ' Total header formatting
         .range("O2:O4").Font.Bold = True
         .range("P1").Font.Bold = True
         .range("Q1").Font.Bold = True
         ' Resulting cells formatting
         .range("P2:Q4").Interior.ColorIndex = 36
         
        ' Find the last row that contains data
        l_RowLast = .Cells(Rows.Count, 1).End(xlUp).Row
        ' Find the last column that contains data
        l_ColLast = .Cells(1, Columns.Count).End(xlToLeft).Column
        ' Find the last column letter
        s_ColLastLetter = Split((.Columns(l_ColLast).Address(, 0)), ":")(0)

        ' change / adjust the size of resulting array
        j = 1
        ReDim arr_results(1 To j)
        
        ' Assignment for the first iteration
        arr_results(j).ticker = .Cells(2, 1).Value
        arr_results(j).price_begin = .Cells(2, 3).Value
        arr_results(j).volume = .Cells(2, 7).Value
        
        ' setup a loop that will go until it reaches the last row
        For i = 3 To l_RowLast
            If arr_results(j).ticker = .Cells(i, 1).Value Then
                ' Summarise volume
                arr_results(j).volume = arr_results(j).volume + .Cells(i, 7).Value
            Else
                ' Close price for previous ticker
                arr_results(j).price_end = .Cells(i - 1, 6).Value

               ' change / adjust the size of array
               j = j + 1
               ReDim Preserve arr_results(1 To j)
               
                ' Assignments for new ticker
                arr_results(j).ticker = .Cells(i, 1).Value
                arr_results(j).price_begin = .Cells(i, 3).Value
                arr_results(j).volume = .Cells(i, 7).Value
            End If
        Next i
        ' Close price for last ticker
        arr_results(j).price_end = .Cells(l_RowLast, 6).Value

        ' Print total data for each ticker under header row
        For j = 1 To UBound(arr_results)
            ' change of price
            d_price_change = arr_results(j).price_end - arr_results(j).price_begin
            
            .Cells(j + 1, 9).Value = arr_results(j).ticker
            .Cells(j + 1, 10).Value = d_price_change
            If arr_results(j).price_begin = 0 Then
                .Cells(j + 1, 11).Value = 0
            Else
                .Cells(j + 1, 11).Value = Application.WorksheetFunction.Round(d_price_change / arr_results(j).price_begin, 4)
            End If
            .Cells(j + 1, 12).Value = arr_results(j).volume
            
            ' Format cells
            If (.Cells(j + 1, 10).Value >= 0) Then
                .Cells(j + 1, 10).Interior.ColorIndex = 4
            Else
                .Cells(j + 1, 10).Interior.ColorIndex = 3
            End If

        Next j

'        set range for searching of max and min value of percent
        Set r = range(.Cells(2, 11).Address(), .Cells(UBound(arr_results) + 1, 11).Address())
        r.NumberFormat = "0.00%"
        
        ' Greatest % Increase
        v_max = WorksheetFunction.max(r)
        l_Row = WorksheetFunction.Match(v_max, r, 0)
        .range("P2").Value = .Cells(l_Row + 1, 9).Value
        .range("Q2").NumberFormat = "0.00%"
        .range("Q2").Value = v_max

        ' Greatest % Decrease
        v_min = WorksheetFunction.Min(r)
        l_Row = WorksheetFunction.Match(v_min, r, 0)
        .range("P3").Value = .Cells(l_Row + 1, 9).Value
        .range("Q3").NumberFormat = "0.00%"
        .range("Q3").Value = v_min
        
        ' Set range for searching of max value of volume
        Set r = .range(.Cells(2, 12).Address(), .Cells(UBound(arr_results) + 1, 12).Address())
        r.NumberFormat = "General"
        
        ' Greatest Total Volume
        v_max_vol = WorksheetFunction.max(r)
        l_Row = WorksheetFunction.Match(v_max_vol, r, 0)
        .range("P4").Value = .Cells(l_Row + 1, 9).Value
        .range("Q4").NumberFormat = "General"
        .range("Q4").Value = v_max_vol
        
        ' Set AutoFit
        .Columns("A:" & s_ColLastLetter).AutoFit
        
    End With
    
    Erase arr_results
    Set r = Nothing

End Sub

