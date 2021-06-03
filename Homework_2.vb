
Sub Ticker_total()
    'Initate loop for worksheet loop
    For Each ws In Worksheets
        'creating new output columns
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Open"
        ws.Range("L1").Value = "Close"
        ws.Range("M1").Value = "Change"
        ws.Range("N1").Value = "% Differnece"
        ws.Range("O1").Value = "Volume"


        'Defining the starting row
        Dim srow As Integer
        srow = 2

        'Defining the last row variable
        Dim lrow As Long
        
        'Calculating the last row for the data
        lrow = Cells(Rows.Count, 1).End(xlUp).Row

        'defining each of the variables by header for input
        Dim Ticker As String
        Dim s_date As Integer
        Dim O_Price As Double
        Dim High As Double
        Dim Low As Double
        Dim C_price As Double
        Dim Vol As Double
        Dim percent As Double
        Dim Change As Double

        'Intializing the values for the calculated variables
        Vol = 0
        O_Price = 0
        C_price = 0
        High = 0
        Low = 0

        'Starting Loop for to add values for calcuated variables
        For i = 2 To lrow
            'Checks the ticker symbol for change
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                'setting the opening
                O_Price = ws.Cells(i, 3).Value
                'iterating throug rows in J column for opening
                ws.Range("K" & srow).Value = O_Price
                srow = srow

                'Checks the ticker symbol for change
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Update Ticker
                    Ticker = ws.Cells(i, 1).Value
                    'Printing Value to new column
                    ws.Range("J" & srow).Value = Ticker
                    'Update C_Price
                    C_price = ws.Cells(i, 6).Value
                    'Printing Value to new column
                    ws.Range("L" & srow).Value = C_price
                    'Update Total Volume
                    Vol = Vol + ws.Cells(i, 7).Value
                    'Printing Value to new column
                    ws.Range("O" & srow).Value = Vol
                    'Iterated the row to next row
                    srow = srow + 1
                    'Reset the Volume for the next row
                    Vol = 0
                    'Calculating the change
                    Change = C_price - O_Price
                    'Printing Value to new column
                    'srow -1 because line 66 already iterated +1
                    ws.Range("L" & srow - 1).Value = Change
                    'Assigning Color to Cells based on Value
                        If Change > 0 Then
                            ws.Range("M" & srow - 1).Interior.ColorIndex = 4
                        Else
                        ws.Range("M" & srow - 1).Interior.ColorIndex = 3
                        End If
                        'Calculating Percent and Assinging color
                        'Checking for 0 becasue can not divide by 0
                        If O_Price = 0 Then
                            'Needs to defined as sting Else it will not display properly
                            percent = "0"
                        Else
                        percent = (Change / O_Price)
                        'Printing Value to new column
                        ws.Range("N" & srow - 1).Value = percent
                        'Percent Formating
                        'https://stackoverflow.com/questions/41510066/calculating-percentage-using-vba
                        ws.Range("N" & srow - 1).NumberFormat = "0.00%"
                        End If
                    Else
                    Vol = Vol + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub
