Sub testData()
        
'/*Declared Variables*/
    Dim openBalance As Double
    Dim closeBalance As Double
    Dim newOpenBalance As Double
    Dim pOpenBalance As Double
    Dim pChng As Double
    Dim vol As Long
    Dim totalVolume As Double
    Dim totalRows As Long
    Dim rnum, stdrow As Integer
    Dim Ychng As Double
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Dim rTotal As Integer

    
    
    Set starting_ws = ActiveSheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
        '/*Get the total number of Rows*/
            totalRows = Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (totalRows)
            
        '/*Setting Variable to enter Unique Value in Row 9*/
            rnum = 1 '//For Ticker
            stdrow = 1 '//For volume
            
        '/*Enter Headers*/
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
        '/*Reading the first Unique Values*/
            Cells(rnum + 1, 9).Value = Cells(2, 1).Value
        
        '/*Etting Volume to 0*/
            totalVolume = 0
            
        '/*Hardcoding the first unique value*/
            rnum = rnum + 1
            Cells(rnum, 9).Value = Cells(rnum, 1).Value
            
        '/*Hardcoding the opening balance*/
            openBalance = Cells(2, 3).Value
            'MsgBox (openBalance)
            
        '/********************************************************/
        '*For Loop to go over different to get the data fron     *
        '*Columns in the sheet to retrieve Ticker Symbol,        *
        '*Opening Balance, Closing Balance, and Volume to        *
        '*perform respective tasks to get the desired output     *
        '*and displaying the result on the sheet under the       *
        '*headers that were created before in the file. IF Loop  *
        '*checks to where the previous Ticker symbol is different*
        '*from the next Ticker Symbol, and then outputs in onto  *
        '*the sheet. After it outputs the Ticker symbol it gets  *
        '*opening balance and the closing balance, performed the *
        '*required calculations and ouput them on the sheet.     *
        '*Later it gets the Volume and perfrom calculations on it*
        '/********************************************************/
            For i = 2 To totalRows
                
        '/*Checking to seperate the different symbols*/
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
        '/*Incrementing row number*/
                    rnum = rnum + 1
                    stdrow = stdrow + 1
                    
        '/*Adding the unique symbol to the Ticker Column*/
                    Cells(rnum, 9).Value = Cells(i + 1, 1).Value
                    'MsgBox (i)
                    
        '/******************************************************/
        '/*Checking to see the value in openBalance Variable to*
        '*calculation because we hard coded the openBalance to *
        '*get the first opening Balance. If it is not equal to *
        '*the value in the first row then get the next values. *
        '/******************************************************/
                    closeBalance = Cells(i, 6).Value
                    newOpenBalance = Cells(i + 1, 3).Value
                     If (openBalance = Cells(2, 3).Value) Then
                        Ychng = closeBalance - openBalance
                        pChng = Ychng / openBalance
                        'On Error GoTo 0
                        Cells(2, 10).Value = Ychng
                        Cells(2, 11).Value = pChng
                        openBalance = 0
                    ElseIf (pOpenBalance = 0) Then
                        pChng = 0
                        Ychng = 0
                        Cells(rnum - 1, 10).Value = Ychng
                        Cells(rnum - 1, 11).Value = pChng
                    Else
                        Ychng = closeBalance - pOpenBalance
                        pChng = Ychng / pOpenBalance
                        'On Error Resume Next
                        Cells(rnum - 1, 10).Value = Ychng
                        Cells(rnum - 1, 11).Value = pChng
                     End If
                     
                     pOpenBalance = newOpenBalance
                    
        
        '/*****************************************************/
        '*Getting the last value of volume for the previous   *
        '*Ticker symbol, calculating the total, and outputing *
        '*it into the sheet.                                  *
        '/*****************************************************/
                    vol = Cells(i, 7).Value
                    totalVolume = totalVolume + vol
                    Cells(stdrow, 12).Value = totalVolume
                    
        '/******************************************************/
        '*Setting all values to 0 for it to start calculating  *
        '*the volume for the next Ticker Symbol.                   *
        '/******************************************************/
                    totalVolume = 0
                    openBalance = 0
                    closeBalance = 0
                    Ychng = 0
                Else
                
        '/*Getting the volume from the vol field in the sheet*/
                    vol = Cells(i, 7).Value
                    
        '//Calculating to total volume and outputting in onto the sheet*/
                    totalVolume = totalVolume + vol
                    Cells(stdrow + 1, 12).Value = totalVolume
                End If
            Next i
            

            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest & Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            
        '//Setting color for Column J
            rTotal = Cells(Rows.Count, 10).End(xlUp).Row
            For k = 2 To rTotal
                If (Cells(k, 10).Value < 0) Then
                    'MsgBox ("Yes at " & k)
                    Cells(k, 10).Interior.ColorIndex = 3
                ElseIf (Cells(k, 10).Value > 0) Then
                    Cells(k, 10).Interior.ColorIndex = 4
                End If
            Next k
            
            Dim m As Double
            ws.Range("J2:J" & rTotal).NumberFormat = "0.000000000"
            ws.Range("K2:K" & rTotal).NumberFormat = "0.00%"
            Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & rTotal))
            Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & rTotal))
            Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & rTotal))
            For i = 2 To rTotal
                If (Cells(2, 17).Value = Cells(i, 11).Value) Then
                    Cells(2, 17).NumberFormat = "0.00%"
                    Cells(2, 16).Value = Cells(i, 9).Value
                ElseIf (Cells(3, 17).Value = Cells(i, 11).Value) Then
                    Cells(3, 17).NumberFormat = "0.00%"
                    Cells(3, 16).Value = Cells(i, 9).Value
                ElseIf (Cells(4, 17).Value = Cells(i, 12).Value) Then
                    Cells(4, 16).Value = Cells(i, 9).Value
                End If
            Next i
            
            ws.Range("A1:Q1").EntireColumn.AutoFit

    Next
    
    starting_ws.Activate
    
End Sub



