Attribute VB_Name = "mod_14_HARQSplitting"
Option Explicit

' Module Name: HARQSplit
' Status: V1.1.0
' Changes in this version:
'   - HARQ_indicator explicitly kept as integer
'   - Split rows leave HARQ_TX_SFN_offset blank
'   - Split row TX_SFN_est = original TX_SFN_est + HARQ_TX_SFN_offset
'   - Final table sorted by TX_SFN_est after splitting
'   - Preserves TXQTIME on split TX2 row (same as original/TX1 row)
'   - Recomputes RX_COUNT after splitting
'   - Recomputes AVG_TOTAL_LATENCY after splitting using the non-blank Avg TX1/Avg TX2 RXTIMES field
'   - No formula save/restore logic
'   - Returns execution time through ByRef argument
' Target: Excel 2024 LTSC

Public Sub Run_HARQSplit()
    Dim elapsedSeconds As Double
    HARQSplit elapsedSeconds
    MsgBox "HARQ split completed in " & Format(elapsedSeconds, "0.00") & " s", vbInformation, "HARQ Split"
End Sub

Public Sub HARQSplit(ByRef elapsedSeconds As Double)
    Dim t0 As Double
    t0 = Timer
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim wsExp As Worksheet
    Set wsExp = wb.Sheets("ExpResults")
    
    Dim loExp As ListObject
    Set loExp = wsExp.ListObjects("ExpResultsTable")
    
    If loExp.DataBodyRange Is Nothing Then
        elapsedSeconds = TimerDiffSeconds(t0, Timer)
        Exit Sub
    End If
    
    Dim numStations As Long
    numStations = CLng(Evaluate(ThisWorkbook.Names("Num_Rx_Stations").RefersTo))
    
    Dim colHARQOffset As Long
    Dim colHARQIndicator As Long
    Dim colThresh As Long
    Dim colTXQTIME As Long
    Dim colTXSFN As Long
    Dim colRXCOUNT As Long
    Dim colAvgTx1 As Long
    Dim colAvgTx2 As Long
    Dim colAvgTotalLatency As Long
    Dim colMsgGenTime As Long
    Dim colRSSI1 As Long
    
    colHARQOffset = GetTableColumnIndex(loExp, "HARQ_TX_SFN_offset")
    colHARQIndicator = GetTableColumnIndex(loExp, "HARQ_indicator")
    colThresh = GetTableColumnIndex(loExp, "HARQ_RXTIME_threshold")
    colTXQTIME = GetTableColumnIndex(loExp, "TXQTIME")
    colTXSFN = GetTableColumnIndex(loExp, "TX_SFN_est")
    colRXCOUNT = GetTableColumnIndex(loExp, "RX_COUNT")
    colAvgTx1 = GetTableColumnIndex(loExp, "Avg TX1 RXTIMES")
    colAvgTx2 = GetTableColumnIndex(loExp, "Avg TX2 RXTIMES")
    colAvgTotalLatency = GetTableColumnIndex(loExp, "AVG_TOTAL_LATENCY")
    colMsgGenTime = GetTableColumnIndex(loExp, "MSG_GEN_TIME")
    colRSSI1 = GetTableColumnIndex(loExp, "RSSI1")
    
    Dim rxTimeCols() As Long
    rxTimeCols = GetSequentialColumns(loExp, "RXTIME")
    
    If UBound(rxTimeCols) < numStations Then
        Err.Raise vbObjectError + 1200, "HARQSplit", "Not enough RXTIME columns found for Num_Rx_Stations."
    End If
    
    Dim srcData As Variant
    srcData = loExp.DataBodyRange.Value2
    
    Dim rowCount As Long
    Dim colCount As Long
    rowCount = UBound(srcData, 1)
    colCount = UBound(srcData, 2)
    
    Dim outputData() As Variant
    ReDim outputData(1 To rowCount * 2, 1 To colCount)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim outIdx As Long
    Dim addedRows As Long
    Dim progressStep As Long
    
    progressStep = 1000
    If progressStep < 1 Then progressStep = 1
    
    Application.StatusBar = "HARQ split: initializing..."
    DoEvents
    
    For i = 1 To rowCount
        If (i Mod progressStep) = 0 Or i = 1 Or i = rowCount Then
            Application.StatusBar = "HARQ split: row " & i & " / " & rowCount & " (" & Format(100# * i / rowCount, "0.0") & "%)"
            DoEvents
        End If
        
        Dim currentOffset As Long
        Dim currentIndicator As Long
        Dim thresh As Double
        
        currentOffset = 0
        currentIndicator = 0
        thresh = 0#
        
        If IsNumeric(srcData(i, colHARQOffset)) Then currentOffset = CLng(srcData(i, colHARQOffset))
        If IsNumeric(srcData(i, colHARQIndicator)) Then currentIndicator = CLng(srcData(i, colHARQIndicator))
        If IsNumeric(srcData(i, colThresh)) Then thresh = CDbl(srcData(i, colThresh))
        
        ' Always create first/output row
        outIdx = outIdx + 1
        For k = 1 To colCount
            outputData(outIdx, k) = srcData(i, k)
        Next k
        
        If LenB(CStr(outputData(outIdx, colHARQIndicator))) = 0 Then
            outputData(outIdx, colHARQIndicator) = CLng(0)
        ElseIf IsNumeric(outputData(outIdx, colHARQIndicator)) Then
            outputData(outIdx, colHARQIndicator) = CLng(outputData(outIdx, colHARQIndicator))
        End If
        
        ' Split only rows with positive HARQ offset and not already marked as TX2 rows
        If currentOffset > 0 And currentIndicator >= 0 Then
            addedRows = addedRows + 1
            
            ' Create second row as a copy of the source row
            outIdx = outIdx + 1
            For k = 1 To colCount
                outputData(outIdx, k) = srcData(i, k)
            Next k
            
            ' TXQTIME on split TX2 row remains identical to original/TX1 row
            outputData(outIdx, colTXQTIME) = outputData(outIdx - 1, colTXQTIME)
            
            ' Split row TX_SFN_est = original TX_SFN_est + HARQ_TX_SFN_offset
            If IsNumeric(outputData(outIdx - 1, colTXSFN)) Then
                outputData(outIdx, colTXSFN) = CLng(outputData(outIdx - 1, colTXSFN)) + currentOffset
            Else
                outputData(outIdx, colTXSFN) = vbNullString
            End If
            
            ' HARQ_TX_SFN_offset should NOT be replicated into split row
            outputData(outIdx, colHARQOffset) = vbNullString
            
            ' Partition RXTIME/RSSI values using raw RXTIME threshold
            For j = 1 To numStations
                Dim rxCol As Long
                Dim rssiCol As Long
                Dim rxVal As Variant
                
                rxCol = rxTimeCols(j)
                rssiCol = colRSSI1 + (j - 1)
                rxVal = outputData(outIdx - 1, rxCol)
                
                If IsNumeric(rxVal) Then
                    If CDbl(rxVal) <> 0# Then
                        If CDbl(rxVal) > thresh Then
                            ' Belongs to TX2 row; clear from TX1 row
                            outputData(outIdx - 1, rxCol) = vbNullString
                            If rssiCol <= colCount Then outputData(outIdx - 1, rssiCol) = vbNullString
                        Else
                            ' Belongs to TX1 row; clear from TX2 row
                            outputData(outIdx, rxCol) = vbNullString
                            If rssiCol <= colCount Then outputData(outIdx, rssiCol) = vbNullString
                        End If
                    End If
                Else
                    ' If not numeric in source row, clear from both split sides for safety consistency
                    outputData(outIdx - 1, rxCol) = vbNullString
                    outputData(outIdx, rxCol) = vbNullString
                    If rssiCol <= colCount Then
                        outputData(outIdx - 1, rssiCol) = vbNullString
                        outputData(outIdx, rssiCol) = vbNullString
                    End If
                End If
            Next j
            
            ' Set HARQ indicators on split pair
            outputData(outIdx - 1, colHARQIndicator) = CLng(currentOffset)
            outputData(outIdx, colHARQIndicator) = CLng(-currentOffset)
            
            ' Keep row-specific Avg TX fields consistent with split role
            outputData(outIdx - 1, colAvgTx2) = vbNullString
            outputData(outIdx, colAvgTx1) = vbNullString
            
            ' Recompute RX_COUNT and AVG_TOTAL_LATENCY for split pair rows
            outputData(outIdx - 1, colRXCOUNT) = CountNonBlankRxTimes(outputData, outIdx - 1, rxTimeCols, numStations)
            outputData(outIdx, colRXCOUNT) = CountNonBlankRxTimes(outputData, outIdx, rxTimeCols, numStations)
            
            outputData(outIdx - 1, colAvgTotalLatency) = ComputeAvgTotalLatency(outputData, outIdx - 1, colAvgTx1, colAvgTx2, colMsgGenTime)
            outputData(outIdx, colAvgTotalLatency) = ComputeAvgTotalLatency(outputData, outIdx, colAvgTx1, colAvgTx2, colMsgGenTime)
        Else
            ' Recompute RX_COUNT and AVG_TOTAL_LATENCY for unchanged row as written
            outputData(outIdx, colRXCOUNT) = CountNonBlankRxTimes(outputData, outIdx, rxTimeCols, numStations)
            outputData(outIdx, colAvgTotalLatency) = ComputeAvgTotalLatency(outputData, outIdx, colAvgTx1, colAvgTx2, colMsgGenTime)
        End If
    Next i
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    loExp.DataBodyRange.Delete
    loExp.Resize loExp.Range.Resize(outIdx + 1, colCount)
    loExp.DataBodyRange.Value2 = outputData
    
    ' HARQ_indicator integer formatting
    On Error Resume Next
    loExp.ListColumns(colHARQIndicator).DataBodyRange.NumberFormat = "0"
    On Error GoTo 0
    
    ' Final sort by TX_SFN_est
    On Error Resume Next
    If colTXSFN > 0 Then
        With loExp.Sort
            .SortFields.Clear
            .SortFields.Add key:=loExp.ListColumns(colTXSFN).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    elapsedSeconds = TimerDiffSeconds(t0, Timer)
End Sub

Private Function ComputeAvgTotalLatency(ByRef arr As Variant, ByVal rowNum As Long, ByVal colAvgTx1 As Long, ByVal colAvgTx2 As Long, ByVal colMsgGenTime As Long) As Variant
    Dim avgRx As Double
    Dim msgGen As Double
    
    If Not IsNumeric(arr(rowNum, colMsgGenTime)) Then
        ComputeAvgTotalLatency = vbNullString
        Exit Function
    End If
    
    msgGen = CDbl(arr(rowNum, colMsgGenTime))
    
    If IsNumeric(arr(rowNum, colAvgTx1)) And LenB(CStr(arr(rowNum, colAvgTx1))) > 0 Then
        avgRx = CDbl(arr(rowNum, colAvgTx1))
        ComputeAvgTotalLatency = avgRx - msgGen
        Exit Function
    End If
    
    If IsNumeric(arr(rowNum, colAvgTx2)) And LenB(CStr(arr(rowNum, colAvgTx2))) > 0 Then
        avgRx = CDbl(arr(rowNum, colAvgTx2))
        ComputeAvgTotalLatency = avgRx - msgGen
        Exit Function
    End If
    
    ComputeAvgTotalLatency = vbNullString
End Function

Private Function CountNonBlankRxTimes(ByRef arr As Variant, ByVal rowNum As Long, ByRef rxTimeCols() As Long, ByVal numStations As Long) As Long
    Dim j As Long
    Dim cnt As Long
    
    For j = 1 To numStations
        If rxTimeCols(j) > 0 Then
            If LenB(CStr(arr(rowNum, rxTimeCols(j)))) > 0 Then
                cnt = cnt + 1
            End If
        End If
    Next j
    
    CountNonBlankRxTimes = cnt
End Function

Private Function GetTableColumnIndex(ByVal lo As ListObject, ByVal headerText As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If Trim(CStr(lc.Name)) = headerText Then
            GetTableColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    Err.Raise vbObjectError + 1201, "HARQSplit", "Required column not found: " & headerText
End Function

Private Function GetSequentialColumns(ByVal lo As ListObject, ByVal prefix As String) As Long()
    Dim maxN As Long
    Dim lc As ListColumn
    Dim suffix As String
    Dim n As Long
    
    For Each lc In lo.ListColumns
        If UCase$(Left$(Trim$(lc.Name), Len(prefix))) = UCase$(prefix) Then
            suffix = Mid$(Trim$(lc.Name), Len(prefix) + 1)
            If IsNumeric(suffix) Then
                If CLng(suffix) > maxN Then maxN = CLng(suffix)
            End If
        End If
    Next lc
    
    Dim arr() As Long
    ReDim arr(1 To maxN)
    
    For Each lc In lo.ListColumns
        If UCase$(Left$(Trim$(lc.Name), Len(prefix))) = UCase$(prefix) Then
            suffix = Mid$(Trim$(lc.Name), Len(prefix) + 1)
            If IsNumeric(suffix) Then
                n = CLng(suffix)
                arr(n) = lc.Index
            End If
        End If
    Next lc
    
    GetSequentialColumns = arr
End Function

Private Function TimerDiffSeconds(ByVal tStart As Double, ByVal tEnd As Double) As Double
    If tEnd >= tStart Then
        TimerDiffSeconds = tEnd - tStart
    Else
        TimerDiffSeconds = (86400# - tStart) + tEnd
    End If
End Function
