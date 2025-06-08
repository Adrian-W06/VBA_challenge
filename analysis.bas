Sub AnalyzeStocksWithOverallSummary()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim volumeTotal As Double
    Dim summaryRow As Long
    Dim outputCol As Integer

    ' Variables to track per sheet maxes
    Dim maxIncrease As Double, minIncrease As Double, maxVolume As Double
    Dim maxTicker As String, minTicker As String, volTicker As String
    Dim maxSheet As String, minSheet As String, volSheet As String

    ' Variables to track overall maxes
    Dim overallMaxIncrease As Double: overallMaxIncrease = -999999
    Dim overallMinIncrease As Double: overallMinIncrease = 999999
    Dim overallMaxVolume As Double: overallMaxVolume = -999999
    Dim overallMaxTicker As String, overallMinTicker As String, overallVolTicker As String
    Dim overallMaxSheet As String, overallMinSheet As String, overallVolSheet As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Overall Summary" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Output header to the right (starts in column I)
            outputCol = 9 ' Column I
            ws.Cells(1, outputCol).Resize(1, 4).Value = Array("Ticker", "Quarterly Change", "Percentage Change", "Total Volume")
            summaryRow = 2

            ' Reset per-sheet trackers
            maxIncrease = -999999
            minIncrease = 999999
            maxVolume = -999999

            For i = 2 To lastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    openPrice = ws.Cells(i, 3).Value ' Open in column C
                    volumeTotal = 0
                End If

                volumeTotal = volumeTotal + ws.Cells(i, 7).Value ' Volume in column G

                If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                    closePrice = ws.Cells(i, 6).Value ' Close in column F

                    Dim change As Double
                    Dim pctChange As Double

                    change = closePrice - openPrice
                    If openPrice <> 0 Then
                        pctChange = (change / openPrice) * 100
                    Else
                        pctChange = 0
                    End If

                    ' Write to the right
                    ws.Cells(summaryRow, outputCol).Value = ticker
                    ws.Cells(summaryRow, outputCol + 1).Value = Round(change, 2)
                    ws.Cells(summaryRow, outputCol + 2).Value = Round(pctChange, 2) & "%"
                    ws.Cells(summaryRow, outputCol + 3).Value = volumeTotal

                    ' Update per-sheet maxes
                    If pctChange > maxIncrease Then
                        maxIncrease = pctChange: maxTicker = ticker: maxSheet = ws.Name
                    End If
                    If pctChange < minIncrease Then
                        minIncrease = pctChange: minTicker = ticker: minSheet = ws.Name
                    End If
                    If volumeTotal > maxVolume Then
                        maxVolume = volumeTotal: volTicker = ticker: volSheet = ws.Name
                    End If

                    ' Update overall maxes
                    If pctChange > overallMaxIncrease Then
                        overallMaxIncrease = pctChange: overallMaxTicker = ticker: overallMaxSheet = ws.Name
                    End If
                    If pctChange < overallMinIncrease Then
                        overallMinIncrease = pctChange: overallMinTicker = ticker: overallMinSheet = ws.Name
                    End If
                    If volumeTotal > overallMaxVolume Then
                        overallMaxVolume = volumeTotal: overallVolTicker = ticker: overallVolSheet = ws.Name
                    End If

                    summaryRow = summaryRow + 1
                End If
            Next i

            ' Write per-sheet summary
            Dim labelRow As Long
            labelRow = summaryRow + 2
            ws.Cells(labelRow, outputCol).Value = "Greatest % Increase:"
            ws.Cells(labelRow, outputCol + 1).Value = maxTicker
            ws.Cells(labelRow, outputCol + 2).Value = Round(maxIncrease, 2) & "%"

            ws.Cells(labelRow + 1, outputCol).Value = "Greatest % Decrease:"
            ws.Cells(labelRow + 1, outputCol + 1).Value = minTicker
            ws.Cells(labelRow + 1, outputCol + 2).Value = Round(minIncrease, 2) & "%"

            ws.Cells(labelRow + 2, outputCol).Value = "Greatest Total Volume:"
            ws.Cells(labelRow + 2, outputCol + 1).Value = volTicker
            ws.Cells(labelRow + 2, outputCol + 3).Value = maxVolume
        End If
    Next ws

    ' Create or clear "Overall Summary" sheet
    Dim summaryWS As Worksheet
    On Error Resume Next
    Set summaryWS = Sheets("Overall Summary")
    If summaryWS Is Nothing Then
        Set summaryWS = Sheets.Add(After:=Sheets(Sheets.Count))
        summaryWS.Name = "Overall Summary"
    Else
        summaryWS.Cells.Clear
    End If
    On Error GoTo 0

    ' Output overall summary values
    With summaryWS
        .Range("A1").Value = "Overall Summary"
        .Range("A3").Value = "Greatest % Increase:"
        .Range("B3").Value = overallMaxTicker
        .Range("C3").Value = overallMaxSheet
        .Range("D3").Value = Round(overallMaxIncrease, 2) & "%"

        .Range("A4").Value = "Greatest % Decrease:"
        .Range("B4").Value = overallMinTicker
        .Range("C4").Value = overallMinSheet
        .Range("D4").Value = Round(overallMinIncrease, 2) & "%"

        .Range("A5").Value = "Greatest Total Volume:"
        .Range("B5").Value = overallVolTicker
        .Range("C5").Value = overallVolSheet
        .Range("D5").Value = overallMaxVolume
    End With

    MsgBox "All summaries completed, including Overall Summary sheet.", vbInformation
End Sub
