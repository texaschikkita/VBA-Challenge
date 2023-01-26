Attribute VB_Name = "Module1"
Sub StockMarket_Analysis()

' Define everything:

' Ticker as string
Dim Ticker As String

' Define variables for: year_open, year_close, yearly_change, total_stock_volume, percent_change
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double

' Define a variable to set up a row to start
Dim start_data As Integer

' Define variable of the worksheet to excute the code in all work sheet at once in the workbook
Dim ws As Worksheet

' Loop through all worksheets.  Excute the code once
For Each ws In Worksheets

    ' Assign headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N1").Value = "Statistics"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
    Range("N1").EntireColumn.AutoFit
    Range("P1").EntireColumn.AutoFit
    Range("J1").EntireColumn.AutoFit
    Range("O1").EntireColumn.AutoFit
    Range("I1").EntireColumn.AutoFit
    Range("K1").EntireColumn.AutoFit
    Range("L1").EntireColumn.AutoFit
    Range("A1").EntireRow.HorizontalAlignment = xlCenter
    Range("A1").EntireRow.Font.Bold = True
    Range("N1").EntireColumn.Font.Bold = True
    Range("J1").EntireColumn.Font.Bold = True
    
    
    
    
    ' Assign starting integer
    start_data = 2
    previous_i = 1
    Total_Stock_Volume = 0

    ' Define last row (last row of column A)
    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        ' For loop to go through all data to find: yearly change, percent change and total stock volume
        For i = 2 To EndRow

            ' If statement to identify same ticker sym or new ticker sym; If same add to vote count; If different create new entry to tally
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Get the Ticker symbol and append to line for dara analysis on that ticker
            Ticker = ws.Cells(i, 1).Value

            ' Variable to go to next Ticker
            previous_i = previous_i + 1

            ' Set range for opening and closing values
            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value

            ' For loop to calculate the total stock volume
            For j = previous_i To i
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j
            

            ' When loop finishes with ticker, reset value to 0 to start again with next unique ticker
            If year_open = 0 Then

                Percent_Change = year_close

            Else
                Yearly_Change = year_close - year_open
                Percent_Change = Yearly_Change / year_open

            End If
        

            ' Add to summary table row
            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change

            ' Calculate for percentage (%) format
            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume


            ' When done with first row, go to next row of summary table
            start_data = start_data + 1

            ' After each ticker with the same name, reset variable to zero
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

            'Move i number to variable previous_i
            previous_i = i

        End If

    ' End of loop

    Next i

' Define second summery table
    ' Go to the last row of column k
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    ' Define variables to initiate the second summery table value
    Increase = 0
    Decrease = 0
    Greatest = 0

        ' Loop to find max/min for percentage change and max volume
        For k = 3 To kEndRow

            ' Define previous increment to check
            last_k = k - 1

            ' Define current row for percentage
            current_k = ws.Cells(k, 11).Value

            ' Define previous row for percentage
            prevous_k = ws.Cells(last_k, 11).Value

            ' Specify where greatest total volume goes in summary table row
            volume = ws.Cells(k, 12).Value

            ' Prevous greatest volume row
            prevous_vol = ws.Cells(last_k, 12).Value

            ' If statement to find increase
            If Increase > current_k And Increase > prevous_k Then
                Increase = Increase

                ' Define name for increase percentage
                ' Increase_name = ws.Cells(k, 9).Value

            ElseIf current_k > Increase And current_k > prevous_k Then
                Increase = current_k

                ' Define name for increase percentage
                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then
                Increase = prevous_k

                ' Define name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value

            End If

       
            'If statement to find the decrease
            If Decrease < current_k And Decrease < prevous_k Then
                'Define decrease as decrease
                Decrease = Decrease

            ' Define name for increase percentage
            ElseIf current_k < Increase And current_k < prevous_k Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then
                Decrease = prevous_k
                decrease_name = ws.Cells(last_k, 9).Value

            End If

           ' If statement to find greatest volume
            If Greatest > volume And Greatest > prevous_vol Then
                Greatest = Greatest
                ' Define name for greatest volume
                'greatest_name = ws.Cells(k, 9(.Value

            ElseIf volume > Greatest And volume > prevous_vol Then
                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                Greatest = prevous_vol
                greatest_name = ws.Cells(last_k, 9).Value

            End If

        Next k

    ' Assign names for greatest increase,greatest decrease, and  greatest volume
    ws.Range("N1").Value = "Statistics"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"
   

    'Get values for greatest increase, greatest increase and ticker with greatest volume
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest

    ' Greatest increase and decrease in percentage format
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    
    ' Apply conditional formatting for end row in column J
    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        ' For loop to specify conditional formatting parameters
        For j = 2 To jEndRow

            ' If greater than or less than zero
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

' Move on to next worksheet
Next ws

End Sub


