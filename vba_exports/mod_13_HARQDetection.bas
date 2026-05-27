Attribute VB_Name = "mod_13_HARQDetection"
Option Explicit

' Module Name: HARQDetection
' Status: V1.0.2
' Changes in this version:
'   - Fixed HARQ detection to classify single-cluster rows using aligned_RXTIMEs first
'   - Single-cluster rows now resolve directly from aligned weighted mean:
'       * near 0 => TX1-only => HARQ_TX_SFN_offset = 0
'       * above minHARQoffset/2 => TX2-only => HARQ_TX_SFN_offset = rounded aligned mean
'   - Two-cluster rows continue to use bounded discrete WLS search
'   - HARQ_TX_SFN_offset is written as integer-valued output and formatted as "0"
'   - Retains progress bar/status bar behavior
' Target: Excel 2024 LTSC

Public Sub Run_HARQDetection()
    Dim elapsedSeconds As Double
    HARQDetection elapsedSeconds
    MsgBox "HARQ detection completed in " & Format(elapsedSeconds, "0.00") & " s", vbInformation, "HARQ Detection"
End Sub

Public Sub HARQDetection(ByRef elapsedSeconds As Double)
    Dim t0 As Double
    t0 = Timer
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim wsExp As Worksheet
    Dim wsCfg As Worksheet
    Dim wsPdu As Worksheet
    
    Set wsExp = wb.Sheets("ExpResults")
    Set wsCfg = wb.Sheets("Exp Config & Data Proc Params")
    Set wsPdu = wb.Sheets("PDU Size Table")
    
    Dim loExp As ListObject
    Dim loStationVendor As ListObject
    Dim loPduRx As ListObject
    Dim loAduMap As ListObject
    
    Set loExp = wsExp.ListObjects("ExpResultsTable")
    Set loStationVendor = wsCfg.ListObjects("StationID2VendorID")
    Set loPduRx = wsCfg.ListObjects("PDU2RXTprocVendorID")
    Set loAduMap = wsPdu.ListObjects("ADU2NumSubchansTable")
    
    If loExp.DataBodyRange Is Nothing Then
        elapsedSeconds = TimerDiffSeconds(t0, Timer)
        Exit Sub
    End If
    
    Dim minHARQOffset As Long
    Dim maxHARQOffset As Long
    minHARQOffset = CLng(Evaluate(ThisWorkbook.Names("minHARQoffset").RefersTo))
    maxHARQOffset = CLng(Evaluate(ThisWorkbook.Names("maxHARQoffset").RefersTo))
    
    Dim dictStationToVendor As Object
    Dim dictPduVendorToRxMu As Object
    Dim dictPduVendorToRxSigma As Object
    Dim dictAduToPdu As Object
    
    Set dictStationToVendor = CreateObject("Scripting.Dictionary")
    Set dictPduVendorToRxMu = CreateObject("Scripting.Dictionary")
    Set dictPduVendorToRxSigma = CreateObject("Scripting.Dictionary")
    Set dictAduToPdu = CreateObject("Scripting.Dictionary")
    
    BuildStationToVendorMap loStationVendor, dictStationToVendor
    BuildPduVendorRxMaps loPduRx, dictPduVendorToRxMu, dictPduVendorToRxSigma
    BuildAduToPduMap loAduMap, dictAduToPdu
    
    Dim data As Variant
    data = loExp.DataBodyRange.Value2
    
    Dim colLEN As Long, colTXSFN As Long, colRXCOUNT As Long
    Dim colAvgTx1 As Long, colAvgTx2 As Long, colHARQOffset As Long, colHARQThresh As Long
    
    colLEN = GetTableColumnIndex(loExp, "LEN")
    colTXSFN = GetTableColumnIndex(loExp, "TX_SFN_est")
    colRXCOUNT = GetTableColumnIndex(loExp, "RX_COUNT")
    colAvgTx1 = GetTableColumnIndex(loExp, "Avg TX1 RXTIMES")
    colAvgTx2 = GetTableColumnIndex(loExp, "Avg TX2 RXTIMES")
    colHARQOffset = GetTableColumnIndex(loExp, "HARQ_TX_SFN_offset")
    colHARQThresh = GetTableColumnIndex(loExp, "HARQ_RXTIME_threshold")
    
    Dim rxTimeCols() As Long
    rxTimeCols = GetSequentialColumns(loExp, "RXTIME")
    
    Dim rowCount As Long
    rowCount = UBound(data, 1)
    
    Dim r As Long
    Dim progressStep As Long
    progressStep = 1000
    If progressStep < 1 Then progressStep = 1
    
    Dim aduLen As Long
    Dim pduSize As Long
    Dim txSfnEst As Double
    Dim rxCount As Long
    
    Dim alignedVals() As Double
    Dim rawVals() As Double
    Dim sigmas() As Double
    Dim rxIdxs() As Long
    Dim obsCount As Long
    
    Dim harqN As Long
    Dim avgTx1 As Variant
    Dim avgTx2 As Variant
    Dim rxThreshold As Variant
    
    Application.StatusBar = "HARQ detection: initializing..."
    DoEvents
    
    For r = 1 To rowCount
        If (r Mod progressStep) = 0 Or r = 1 Or r = rowCount Then
            Application.StatusBar = "HARQ detection: row " & r & " / " & rowCount & " (" & Format(100# * r / rowCount, "0.0") & "%)"
            DoEvents
        End If
        
        data(r, colAvgTx1) = vbNullString
        data(r, colAvgTx2) = vbNullString
        data(r, colHARQOffset) = vbNullString
        data(r, colHARQThresh) = vbNullString
        
        If IsNumeric(data(r, colLEN)) Then
            aduLen = CLng(data(r, colLEN))
        Else
            GoTo NextRow
        End If
        
        If IsNumeric(data(r, colTXSFN)) Then
            txSfnEst = CDbl(data(r, colTXSFN))
        Else
            GoTo NextRow
        End If
        
        If IsNumeric(data(r, colRXCOUNT)) Then
            rxCount = CLng(data(r, colRXCOUNT))
        Else
            rxCount = 0
        End If
        
        If rxCount <= 0 Then
            data(r, colHARQOffset) = RoundToLong(0)
            GoTo NextRow
        End If
        
        If dictAduToPdu.Exists(CStr(aduLen)) Then
            pduSize = CLng(dictAduToPdu(CStr(aduLen)))
        Else
            pduSize = aduLen
        End If
        
        obsCount = CollectAlignedRxObservations(data, r, rxTimeCols, txSfnEst, pduSize, dictStationToVendor, dictPduVendorToRxMu, dictPduVendorToRxSigma, rawVals, alignedVals, sigmas, rxIdxs)
        
        If obsCount <= 0 Then
            data(r, colHARQOffset) = RoundToLong(0)
            GoTo NextRow
        End If
        
        AnalyzeRowHARQ_V1_0_2 alignedVals, rawVals, sigmas, obsCount, minHARQOffset, maxHARQOffset, harqN, avgTx1, avgTx2, rxThreshold
        
        data(r, colHARQOffset) = RoundToLong(harqN)
        If Not IsEmpty(avgTx1) Then data(r, colAvgTx1) = avgTx1
        If Not IsEmpty(avgTx2) Then data(r, colAvgTx2) = avgTx2
        If Not IsEmpty(rxThreshold) Then data(r, colHARQThresh) = rxThreshold
        
NextRow:
    Next r
    
    loExp.DataBodyRange.Value2 = data
    
    On Error Resume Next
    loExp.ListColumns(colHARQOffset).DataBodyRange.NumberFormat = "0"
    On Error GoTo 0
    
    Application.StatusBar = False
    elapsedSeconds = TimerDiffSeconds(t0, Timer)
End Sub

Private Sub AnalyzeRowHARQ_V1_0_2(ByRef alignedVals() As Double, ByRef rawVals() As Double, ByRef sigmas() As Double, ByVal obsCount As Long, ByVal minHARQOffset As Long, ByVal maxHARQOffset As Long, ByRef harqN As Long, ByRef avgTx1 As Variant, ByRef avgTx2 As Variant, ByRef rxThreshold As Variant)
    Dim i As Long
    Dim spread As Double
    Dim initialN As Long
    Dim wlsMeanAligned As Double
    
    Dim minAligned As Double
    Dim maxAligned As Double
    
    Dim tx1Raw() As Double
    Dim tx1Sig() As Double
    Dim tx2Raw() As Double
    Dim tx2Sig() As Double
    Dim tx1Cnt As Long
    Dim tx2Cnt As Long
    
    avgTx1 = Empty
    avgTx2 = Empty
    rxThreshold = Empty
    harqN = 0
    
    If obsCount <= 0 Then Exit Sub
    
    minAligned = alignedVals(1)
    maxAligned = alignedVals(1)
    For i = 2 To obsCount
        If alignedVals(i) < minAligned Then minAligned = alignedVals(i)
        If alignedVals(i) > maxAligned Then maxAligned = alignedVals(i)
    Next i
    spread = maxAligned - minAligned
    wlsMeanAligned = WeightedMean(alignedVals, sigmas, obsCount)
    
    ' First handle single-cluster rows directly in aligned space.
    ' If the aligned cluster spread is small, classify directly from aligned weighted mean.
    If spread < (minHARQOffset / 2#) Then
        If wlsMeanAligned < (minHARQOffset / 2#) Then
            harqN = 0
            avgTx1 = WeightedMean(rawVals, sigmas, obsCount)
            avgTx2 = Empty
            rxThreshold = MaxArray(rawVals, obsCount) + 1#
        Else
            harqN = RoundToLong(Application.WorksheetFunction.Max(minHARQOffset, wlsMeanAligned))
            If harqN > maxHARQOffset Then harqN = maxHARQOffset
            avgTx1 = Empty
            avgTx2 = WeightedMean(rawVals, sigmas, obsCount)
            rxThreshold = MinArray(rawVals, obsCount) - 1#
        End If
        Exit Sub
    End If
    
    ' Otherwise treat as candidate two-cluster row and use bounded discrete WLS search.
    initialN = RoundToLong(spread)
    If initialN < minHARQOffset Then initialN = minHARQOffset
    If initialN > maxHARQOffset Then initialN = maxHARQOffset
    
    harqN = FindBestHARQOffset(alignedVals, sigmas, obsCount, initialN, minHARQOffset, maxHARQOffset)
    
    ReDim tx1Raw(1 To obsCount)
    ReDim tx1Sig(1 To obsCount)
    ReDim tx2Raw(1 To obsCount)
    ReDim tx2Sig(1 To obsCount)
    
    For i = 1 To obsCount
        If WeightedSqErr(alignedVals(i), 0#, sigmas(i)) <= WeightedSqErr(alignedVals(i), harqN, sigmas(i)) Then
            tx1Cnt = tx1Cnt + 1
            tx1Raw(tx1Cnt) = rawVals(i)
            tx1Sig(tx1Cnt) = sigmas(i)
        Else
            tx2Cnt = tx2Cnt + 1
            tx2Raw(tx2Cnt) = rawVals(i)
            tx2Sig(tx2Cnt) = sigmas(i)
        End If
    Next i
    
    ' If everything falls into one side after search, resolve using aligned weighted mean.
    If tx1Cnt = obsCount Then
        If wlsMeanAligned < (minHARQOffset / 2#) Then
            harqN = 0
            avgTx1 = WeightedMean(rawVals, sigmas, obsCount)
            avgTx2 = Empty
            rxThreshold = MaxArray(rawVals, obsCount) + 1#
        Else
            harqN = RoundToLong(Application.WorksheetFunction.Max(minHARQOffset, wlsMeanAligned))
            If harqN > maxHARQOffset Then harqN = maxHARQOffset
            avgTx1 = Empty
            avgTx2 = WeightedMean(rawVals, sigmas, obsCount)
            rxThreshold = MinArray(rawVals, obsCount) - 1#
        End If
        Exit Sub
    End If
    
    If tx2Cnt = obsCount Then
        harqN = RoundToLong(Application.WorksheetFunction.Max(minHARQOffset, wlsMeanAligned))
        If harqN > maxHARQOffset Then harqN = maxHARQOffset
        avgTx1 = Empty
        avgTx2 = WeightedMean(rawVals, sigmas, obsCount)
        rxThreshold = MinArray(rawVals, obsCount) - 1#
        Exit Sub
    End If
    
    If tx1Cnt > 0 Then avgTx1 = WeightedMean(tx1Raw, tx1Sig, tx1Cnt)
    If tx2Cnt > 0 Then avgTx2 = WeightedMean(tx2Raw, tx2Sig, tx2Cnt)
    
    If tx1Cnt > 0 And tx2Cnt > 0 Then
        rxThreshold = (MinArray(tx2Raw, tx2Cnt) + MaxArray(tx1Raw, tx1Cnt)) / 2#
    End If
End Sub

Private Function FindBestHARQOffset(ByRef alignedVals() As Double, ByRef sigmas() As Double, ByVal obsCount As Long, ByVal initialN As Long, ByVal minHARQOffset As Long, ByVal maxHARQOffset As Long) As Long
    Dim bestN As Long
    Dim bestCost As Double
    Dim candCost As Double
    Dim n As Long
    
    bestN = initialN
    bestCost = HARQCost(alignedVals, sigmas, obsCount, initialN)
    
    For n = initialN - 1 To minHARQOffset Step -1
        candCost = HARQCost(alignedVals, sigmas, obsCount, n)
        If candCost <= bestCost Then
            bestCost = candCost
            bestN = n
        Else
            Exit For
        End If
    Next n
    
    For n = initialN + 1 To maxHARQOffset
        candCost = HARQCost(alignedVals, sigmas, obsCount, n)
        If candCost <= bestCost Then
            bestCost = candCost
            bestN = n
        Else
            Exit For
        End If
    Next n
    
    FindBestHARQOffset = bestN
End Function

Private Function HARQCost(ByRef alignedVals() As Double, ByRef sigmas() As Double, ByVal obsCount As Long, ByVal n As Long) As Double
    Dim i As Long
    Dim c As Double
    
    For i = 1 To obsCount
        c = c + Application.WorksheetFunction.Min(WeightedSqErr(alignedVals(i), 0#, sigmas(i)), WeightedSqErr(alignedVals(i), n, sigmas(i)))
    Next i
    
    HARQCost = c
End Function

Private Function WeightedSqErr(ByVal x As Double, ByVal mu As Double, ByVal sigma As Double) As Double
    Dim s As Double
    s = sigma
    If s <= 0# Then s = 1#
    WeightedSqErr = ((x - mu) * (x - mu)) / (s * s)
End Function

Private Function WeightedMean(ByRef vals() As Double, ByRef sigmas() As Double, ByVal n As Long) As Double
    Dim i As Long
    Dim w As Double
    Dim num As Double
    Dim den As Double
    
    For i = 1 To n
        If sigmas(i) > 0# Then
            w = 1# / (sigmas(i) * sigmas(i))
        Else
            w = 1#
        End If
        num = num + (vals(i) * w)
        den = den + w
    Next i
    
    If den <= 0# Then
        WeightedMean = 0#
    Else
        WeightedMean = num / den
    End If
End Function

Private Function CollectAlignedRxObservations(ByRef data As Variant, ByVal rowNum As Long, ByRef rxTimeCols() As Long, ByVal txSfnEst As Double, ByVal pduSize As Long, ByRef dictStationToVendor As Object, ByRef dictPduVendorToRxMu As Object, ByRef dictPduVendorToRxSigma As Object, ByRef rawVals() As Double, ByRef alignedVals() As Double, ByRef sigmas() As Double, ByRef rxIdxs() As Long) As Long
    Dim i As Long
    Dim c As Long
    Dim stationId As Long
    Dim rxVendor As String
    Dim key As String
    Dim rawRx As Double
    Dim mu As Double
    Dim sigma As Double
    
    ReDim rawVals(1 To UBound(rxTimeCols))
    ReDim alignedVals(1 To UBound(rxTimeCols))
    ReDim sigmas(1 To UBound(rxTimeCols))
    ReDim rxIdxs(1 To UBound(rxTimeCols))
    
    For i = 1 To UBound(rxTimeCols)
        If rxTimeCols(i) > 0 Then
            If IsNumeric(data(rowNum, rxTimeCols(i))) Then
                rawRx = CDbl(data(rowNum, rxTimeCols(i)))
                If rawRx > 0# Then
                    stationId = i
                    If dictStationToVendor.Exists(CStr(stationId)) Then
                        rxVendor = CStr(dictStationToVendor(CStr(stationId)))
                        key = CStr(pduSize) & "|" & CStr(rxVendor)
                        
                        If dictPduVendorToRxMu.Exists(key) Then
                            mu = CDbl(dictPduVendorToRxMu(key))
                        Else
                            mu = 0#
                        End If
                        
                        If dictPduVendorToRxSigma.Exists(key) Then
                            sigma = CDbl(dictPduVendorToRxSigma(key))
                        Else
                            sigma = 1#
                        End If
                        
                        c = c + 1
                        rawVals(c) = rawRx
                        alignedVals(c) = rawRx - txSfnEst - mu
                        sigmas(c) = sigma
                        rxIdxs(c) = i
                    End If
                End If
            End If
        End If
    Next i
    
    If c = 0 Then
        CollectAlignedRxObservations = 0
    Else
        ReDim Preserve rawVals(1 To c)
        ReDim Preserve alignedVals(1 To c)
        ReDim Preserve sigmas(1 To c)
        ReDim Preserve rxIdxs(1 To c)
        CollectAlignedRxObservations = c
    End If
End Function

Private Sub BuildStationToVendorMap(ByVal lo As ListObject, ByRef dictOut As Object)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    Dim arr As Variant
    arr = lo.DataBodyRange.Value2
    
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Trim(CStr(arr(r, 1))) <> "" Then
            dictOut(Trim(CStr(arr(r, 1)))) = Trim(CStr(arr(r, 2)))
        End If
    Next r
End Sub

Private Sub BuildPduVendorRxMaps(ByVal lo As ListObject, ByRef dictMu As Object, ByRef dictSigma As Object)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    Dim lastCol As Long
    lastCol = lo.ListColumns.count
    
    Dim vendorCols As Object
    Set vendorCols = CreateObject("Scripting.Dictionary")
    
    Dim c As Long
    For c = 2 To lastCol Step 2
        If c + 1 <= lastCol Then
            vendorCols(CStr((c \ 2))) = Array(c, c + 1)
        End If
    Next c
    
    Dim arr As Variant
    arr = lo.DataBodyRange.Value2
    
    Dim r As Long, vendorID As Variant, cols As Variant, key As String
    For r = 1 To UBound(arr, 1)
        For Each vendorID In vendorCols.Keys
            cols = vendorCols(vendorID)
            key = CStr(arr(r, 1)) & "|" & CStr(vendorID)
            dictMu(key) = arr(r, cols(0))
            dictSigma(key) = arr(r, cols(1))
        Next vendorID
    Next r
End Sub

Private Sub BuildAduToPduMap(ByVal lo As ListObject, ByRef dictOut As Object)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    Dim arr As Variant
    arr = lo.DataBodyRange.Value2
    
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Trim(CStr(arr(r, 1))) <> "" Then
            If IsNumeric(arr(r, 3)) Then
                dictOut(CStr(arr(r, 1))) = CLng(arr(r, 3))
            End If
        End If
    Next r
End Sub

Private Function GetTableColumnIndex(ByVal lo As ListObject, ByVal headerText As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If Trim(CStr(lc.Name)) = headerText Then
            GetTableColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    Err.Raise vbObjectError + 1000, "HARQDetection", "Required column not found: " & headerText
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

Private Function MaxArray(ByRef vals() As Double, ByVal n As Long) As Double
    Dim i As Long
    Dim m As Double
    m = vals(1)
    For i = 2 To n
        If vals(i) > m Then m = vals(i)
    Next i
    MaxArray = m
End Function

Private Function MinArray(ByRef vals() As Double, ByVal n As Long) As Double
    Dim i As Long
    Dim m As Double
    m = vals(1)
    For i = 2 To n
        If vals(i) < m Then m = vals(i)
    Next i
    MinArray = m
End Function

Private Function RoundToLong(ByVal x As Double) As Long
    RoundToLong = CLng(Application.WorksheetFunction.Round(x, 0))
End Function

Private Function TimerDiffSeconds(ByVal tStart As Double, ByVal tEnd As Double) As Double
    If tEnd >= tStart Then
        TimerDiffSeconds = tEnd - tStart
    Else
        TimerDiffSeconds = (86400# - tStart) + tEnd
    End If
End Function
