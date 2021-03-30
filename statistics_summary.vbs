Sub StatisSummary():
    ' Find out the last used row in the temp
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ' Find out the first and last day
    Dim firstDay As Long
    firstDay = ActiveSheet.Cells(2, 2).Value
    Dim LastDay As Long
    LastDay = ActiveSheet.Cells(lastRow, 2).Value
    ' MsgBox("fisrt day: " & firstDay & "  last day: " & LastDay)

    Dim LastTicker As String
    Dim CrtTicker As String
    Dim NewTickerLocation As Long
    Dim Vol_cum AS DOUBLE
    Dim CRT_vol AS DOUBLE
    LastTicker = ""
    NewTickerLocation = 1
    Vol_cum = 0

    For j = 2 To lastRow
        CrtTicker = ActiveSheet.Range("A" & j).Value
        If CrtTicker = LastTicker Then
            CRT_vol = ActiveSheet.Range("G" & j).value
            ' Msgbox(CRT_vol)
            Vol_cum = Vol_cum + CRT_vol
        ELSE
            LastTicker = ActiveSheet.Range("A" & j).Value
            NewTickerLocation = NewTickerLocation + 1
            ActiveSheet.Range("I" & NewTickerLocation).Value = CrtTicker
            ' Assign open at the first day for current Ticker
            ActiveSheet.Range("M" & NewTickerLocation).Value = ActiveSheet.Range("C" & j).value
            ' Assign close at the last day for the last ticker
            ActiveSheet.Range("N" & (NewTickerLocation - 1)).Value = ActiveSheet.Range("F" & (j - 1)).value
            ' Assign cummulative volumn
            ActiveSheet.Range("L" & (NewTickerLocation - 1)).Value = Vol_cum
            Vol_cum = 0
        End If
    Next j
    ActiveSheet.Range("N" & (NewTickerLocation)).Value = ActiveSheet.Range("F" & lastRow).value
    ActiveSheet.Range("L" & (NewTickerLocation)).Value = Vol_cum
    
    Dim lastRowStat As Long
    lastRowStat = ActiveSheet.Cells(Rows.Count, 13).End(xlUp).Row

    For j = 2 To lastRowStat
        ActiveSheet.Range("J" & j).Value = ActiveSheet.Range("N" & j).Value - ActiveSheet.Range("M" & j).Value
        IF ActiveSheet.Range("M" & j).Value = 0 Then
            ActiveSheet.Range("K" & j).Value = 0
        ELSE
            ActiveSheet.Range("K" & j).Value = ActiveSheet.Range("J" & j).Value / ActiveSheet.Range("M" & j).Value
            IF ActiveSheet.Range("J" & j).Value > 0 Then
                ActiveSheet.Range("J" & j).Interior.ColorIndex = 4
            ELSE
                ActiveSheet.Range("J" & j).Interior.ColorIndex = 3
            END IF
        END IF
        ActiveSheet.Range("K" & j).NumberFormat="0.00%"

    Next j

    ActiveSheet.Range("I1").Value = "<ticker>"
    ActiveSheet.Range("J1").Value = "Yearly Change"
    ActiveSheet.Range("K1").Value = "Percent Change"
    ActiveSheet.Range("L1").Value = "Total Stock Volume"
    ActiveSheet.Range("M1:M" & lastRowStat).ClearContents
    ActiveSheet.Range("N1:N" & lastRowStat).ClearContents
    
    ' BONUS

    Dim lastRowStat2 As Long
    lastRowStat2 = ActiveSheet.Cells(Rows.Count, 9).End(xlUp).Row

    Dim global_inc_max AS DOUBLE
    Dim global_dec_max AS DOUBLE
    Dim global_vol_max AS DOUBLE
    global_inc_max = 0
    global_dec_max = 0
    global_vol_max = 0
    Dim global_inc_max_tic As String
    Dim global_dec_max_tic As String
    Dim global_vol_max_tic As String
    global_inc_max_tic = ""
    global_dec_max_tic = ""
    global_vol_max_tic = ""

    For j = 2 To lastRowStat2
        IF ActiveSheet.Range("K" & j).Value >= global_inc_max Then
            global_inc_max = ActiveSheet.Range("K" & j).Value
            global_inc_max_tic = ActiveSheet.Range("I" & j).Value
            
        END IF
        IF ActiveSheet.Range("K" & j).Value <= global_dec_max Then
            global_dec_max = ActiveSheet.Range("K" & j).Value
            global_dec_max_tic = ActiveSheet.Range("I" & j).Value
        END IF
        IF ActiveSheet.Range("L" & j).Value >= global_vol_max Then
            global_vol_max = ActiveSheet.Range("L" & j).Value
            global_vol_max_tic = ActiveSheet.Range("I" & j).Value
        END IF
    Next j

    ActiveSheet.Range("P2").Value = "Greatest % increase"
    ActiveSheet.Range("Q2").Value = global_inc_max_tic
    ActiveSheet.Range("R2").Value = global_inc_max
    ActiveSheet.Range("R2").NumberFormat="0.00%"

    ActiveSheet.Range("P3").Value = "Greatest % decrease"
    ActiveSheet.Range("Q3").Value = global_dec_max_tic
    ActiveSheet.Range("R3").Value = global_dec_max
    ActiveSheet.Range("R3").NumberFormat="0.00%"

    ActiveSheet.Range("P4").Value = "Greatest total volume"
    ActiveSheet.Range("Q4").Value = global_vol_max_tic
    ActiveSheet.Range("R4").Value = global_vol_max

    ActiveSheet.Range("Q1").Value = "Ticker"
    ActiveSheet.Range("R1").Value = "Value"

    ActiveSheet.UsedRange.EntireColumn.AutoFit
End Sub

