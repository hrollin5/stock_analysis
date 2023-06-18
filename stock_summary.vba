Sub stock_summary():

For Each ws In Worksheets

'define all variables.
Dim i As Long
Dim ticker As String
Dim yearstart As Double
Dim yearend As Double
Dim yearchange As Double
Dim percentchange As Double
Dim volume As LongLong
Dim summaryrow As Integer
Dim lastrow As Long
'find the last row of the data.
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   'set starting value of yearstart
    yearstart = ws.Cells(2, 3).Value
    'labels for summary table and AutoFit
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Range("I1:L1").Columns.AutoFit
    ws.Columns("J").NumberFormat = "$#,##0.00"
    ws.Columns("K").NumberFormat = "0.00%"
    
    'set starting summary table row
    summaryrow = 2
'go through all row of data
For i = 2 To lastrow
    'check if the ticker matches next row, and if it doesn't:
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
           'update variables
            yearend = ws.Cells(i, 6).Value
            yearchange = yearend - yearstart
            percentchange = yearchange / yearstart
            ticker = ws.Cells(i, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            'update summary table
            ws.Range("I" & summaryrow).Value = ticker
            ws.Range("J" & summaryrow).Value = yearchange
            ws.Range("K" & summaryrow).Value = percentchange
            ws.Range("L" & summaryrow).Value = volume
            
            'conditional formatting for summary table
            If ws.Cells(summaryrow, 10).Value < 0 Then
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
            Else: ws.Cells(summaryrow, 10).Interior.ColorIndex = 43
            
            End If
            
            If ws.Cells(summaryrow, 11).Value < 0 Then
                ws.Cells(summaryrow, 11).Interior.ColorIndex = 3
            Else: ws.Cells(summaryrow, 11).Interior.ColorIndex = 43
            
            End If
        
        'reset/update variables
        volume = 0
        percentchange = 0
        yearchange = 0
        yearstart = ws.Cells(i + 1, 3).Value
        summaryrow = summaryrow + 1
   'if the next row is the same ticker
    Else
    volume = volume + ws.Cells(i, 7).Value
    
    End If
    
Next i

'Format Second Table
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

 ws.Cells(2, 16).NumberFormat = "0.00%"
 ws.Cells(3, 16).NumberFormat = "0.00%"
 
'Find last row of summary table
Dim j As Integer
Dim lastrow2 As Integer
lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Set variables to find maximum and minimum percent change and maximum volume
Dim maxpercent As Double
maxpercent = 0
Dim maxticker As String
Dim minpercent As Double
minpercent = 0
Dim minticker As String
Dim maxvolume As LongLong
maxvolume = 0
Dim volticker As String

'Find maxpercent, minpercent, and maxvolume
    For j = 2 To lastrow2
        If ws.Cells(j, 11).Value > maxpercent Then
            maxpercent = ws.Cells(j, 11).Value
            maxticker = ws.Cells(j, 9).Value
        End If
   
        If ws.Cells(j, 11).Value < minpercent Then
          minpercent = ws.Cells(j, 11).Value
           minticker = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 12).Value > maxvolume Then
            maxvolume = ws.Cells(j, 12).Value
            volticker = ws.Cells(j, 9).Value
        End If
    Next j

'insert values into table 2
ws.Range("O2").Value = maxticker
ws.Range("P2").Value = maxpercent

ws.Range("O3").Value = minticker
ws.Range("P3").Value = minpercent

ws.Range("O4").Value = volticker
ws.Range("P4").Value = maxvolume


ws.Range("N:P").Columns.AutoFit

Next ws

End Sub



