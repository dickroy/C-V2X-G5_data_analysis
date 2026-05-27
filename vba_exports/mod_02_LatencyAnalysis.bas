Attribute VB_Name = "mod_02_LatencyAnalysis"
' Module: Latency_Analysis_V2X_LTSC_2024
' Version: 17.3.11-CV2X - VENDOR SEGREGATION INTEGRATED
' Status: PERFORMANCE WRAPPERS, PROGRESS BAR, AND TIMING INTEGRATED

Option Explicit

' Native high-precision timing API (supports 64-bit and 32-bit Excel)
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

' High-precision timer function (returns seconds with microsecond resolution)
Private Function MicroTimer() As Double
    Dim cyTicks As Currency, cyFreq As Currency
    QueryPerformanceFrequency cyFreq
    QueryPerformanceCounter cyTicks
    If cyFreq > 0 Then MicroTimer = cyTicks / cyFreq
End Function

' ===========================================================================================
' WRAPPER FOR MANUAL EXECUTION (Handles UI prompts and performance popup)
' ===========================================================================================
Sub Run_GenerateLatencyAnalysis()
    Dim startTime As Double: startTime = MicroTimer()
    Dim plotIndividual As VbMsgBoxResult
    Dim doIndividualPlots As Boolean
    
    ' Query the user for individual plots (Manual Run Only)
    plotIndividual = MsgBox("Generate individual plots for each Station in addition to Vendor plots?", _
                            vbYesNo + vbQuestion, "Plot Selection")
    
    doIndividualPlots = (plotIndividual = vbYes)
    
    ' Call the core routine with the user's choice (passing Nothing for the logTable)
    Call GenerateLatencyAnalysis(Nothing, doIndividualPlots)
    
    ' Output execution timing only on manual runs
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "Latency Analysis Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT (Safe for automated pipelines calling GenerateLatencyAnalysis(Nothing))
' ===========================================================================================
Sub GenerateLatencyAnalysis(Optional ByRef logTable As Object = Nothing, _
                             Optional ByVal doIndividualPlots As Boolean = False)
    Dim tStart As Double: tStart = MicroTimer() ' Start Timing
    
    Dim wsSrc As Worksheet, wsHist As Worksheet, wsApp As Worksheet, wsPdu As Worksheet
    Dim targetTable As ListObject, appParamsTable As ListObject, pduTable As ListObject, vendorTable As ListObject
    Dim dataArr As Variant, appParamsArr As Variant, pduArr As Variant, vendorArr As Variant, rxCols() As Long, colCount As Long
    Dim txqIdx As Long, genIdx As Long, sfnIdx As Long, txidIdx As Long, msgIdIdx As Long, lenIdx As Long, i As Long
    Dim rowTopAnchor As Double: rowTopAnchor = 20
    
    ' Pre-fetch evaluated values
    Dim minRx As Double, maxRX As Double, stepRX As Double
    Dim minTX As Double, maxTX As Double, stepTX As Double
    Dim minTot As Double, maxTot As Double, stepTot As Double

    On Error Resume Next
    minRx = [MIN_RX_MAC_latency]: maxRX = [MAX_RX_MAC_latency]: stepRX = [BIN_WIDTH_RX_MAC_latency]
    minTX = [MIN_TX_MAC_latency]: maxTX = [MAX_TX_MAC_latency]: stepTX = [BIN_WIDTH_TX_MAC_latency]
    minTot = [MIN_TOT_latency]: maxTot = [MAX_TOT_latency]: stepTot = [BIN_WIDTH_TOT_latency]
    
    ' Error 6 Protection: Ensure no zeros or invalid ranges
    If stepRX <= 0 Then stepRX = 1
    If stepTX <= 0 Then stepTX = 1
    If stepTot <= 0 Then stepTot = 1
    If maxRX <= minRx Then maxRX = minRx + 100
    If maxTX <= minTX Then maxTX = minTX + 100
    If maxTot <= minTot Then maxTot = minTot + 100
    On Error GoTo 0
    
    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    On Error Resume Next
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    On Error GoTo 0
    If targetTable Is Nothing Then
        MsgBox "Error: Table 'ExpResultsTable' not found!", vbCritical
        Exit Sub
    End If

    ' Enable Cover / Performance mode
    CoverScreen True
    UpdateProgressBar 5, "Initializing Data structures..."

    dataArr = targetTable.DataBodyRange.Value
    txqIdx = targetTable.ListColumns("TXQTIME").Index
    genIdx = targetTable.ListColumns("MSG_GEN_TIME").Index
    sfnIdx = targetTable.ListColumns("TX_SFN_est").Index
    txidIdx = targetTable.ListColumns("TX_ID").Index
    lenIdx = targetTable.ListColumns("LEN").Index
    
    ' Locate Message/App ID column for AC Mapping
    On Error Resume Next
    msgIdIdx = targetTable.ListColumns("MSG_ID").Index
    If msgIdIdx = 0 Then msgIdIdx = targetTable.ListColumns("APP_ID").Index
    If msgIdIdx = 0 Then msgIdIdx = targetTable.ListColumns("IVI_ID").Index
    On Error GoTo 0
    
    colCount = 0
    For i = 1 To targetTable.ListColumns.count
        If targetTable.ListColumns(i).Name Like "RXTIME*" Then
            colCount = colCount + 1
            ReDim Preserve rxCols(1 To colCount)
            rxCols(colCount) = i
        End If
    Next i

    ' --- LOAD VENDOR MAPPINGS FROM StationID2VendorID ---
    Dim stToVenMap As Object: Set stToVenMap = CreateObject("Scripting.Dictionary")
    Dim uniqueVendors As Object: Set uniqueVendors = CreateObject("Scripting.Dictionary")
    Set wsApp = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
    
    On Error Resume Next
    Set vendorTable = wsApp.ListObjects("StationID2VendorID")
    On Error GoTo 0
    
    If Not vendorTable Is Nothing Then
        vendorArr = vendorTable.DataBodyRange.Value
        For i = 1 To UBound(vendorArr, 1)
            Dim stId As Variant: stId = vendorArr(i, 1)
            Dim venId As Variant: venId = vendorArr(i, 2)
            If Not IsEmpty(stId) And Not IsEmpty(venId) Then
                stToVenMap(CStr(stId)) = CStr(venId)
                uniqueVendors(CStr(venId)) = True
            End If
        Next i
    End If

    ' --- LOAD AC MAPPINGS FROM APPPARAMS ---
    Dim acMap As Object: Set acMap = CreateObject("Scripting.Dictionary")
    Dim idColParam As Long, acColParam As Long, cleanName As String, key As String
    
    On Error Resume Next
    Set appParamsTable = wsApp.ListObjects("AppParams")
    On Error GoTo 0
    
    If Not appParamsTable Is Nothing Then
        appParamsArr = appParamsTable.DataBodyRange.Value
        For i = 1 To appParamsTable.ListColumns.count
            cleanName = UCase(Trim(appParamsTable.ListColumns(i).Name))
            If cleanName = "MSG_ID" Or cleanName = "APP_ID" Or cleanName = "IVI_ID" Then idColParam = i
            If cleanName = "AC" Or cleanName = "ACCESS_CLASS" Then acColParam = i
        Next i
        
        If idColParam > 0 And acColParam > 0 Then
            For i = 1 To UBound(appParamsArr, 1)
                key = Trim(CStr(appParamsArr(i, idColParam)))
                If key <> "" Then acMap(key) = CInt(appParamsArr(i, acColParam))
            Next i
        End If
    End If

    ' Delete and recreate sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Latency Analysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsHist = ThisWorkbook.Sheets.Add(After:=wsSrc)
    wsHist.Name = "Latency Analysis"

    ' --- EXECUTE WITH PROGRESS TRACKING ---
    UpdateProgressBar 15, "Processing Row 1: RX MAC Latency..."
    ' --- ROW 1: RX MAC LATENCY (Stations) ---
    RunAnalysisBlock wsHist, dataArr, rxCols, sfnIdx, txidIdx, "RX MAC LATENCY", _
                     minRx, maxRX, stepRX, _
                     False, rowTopAnchor, IIf(doIndividualPlots, vbYes, vbNo), False, stToVenMap, uniqueVendors

    ' ===========================================================================================
    ' INJECTED N-ROWS: UNIQUE PDU SIZE (B) FILTERED SUB-ANALYSIS
    ' ===========================================================================================
    On Error Resume Next
    Set wsPdu = ThisWorkbook.Sheets("PDU Size Table")
    Set pduTable = wsPdu.ListObjects("ADU2NumSubchansTable")
    On Error GoTo 0
    
    Dim lenToPduMap As Object: Set lenToPduMap = CreateObject("Scripting.Dictionary")
    Dim lenToMcsMap As Object: Set lenToMcsMap = CreateObject("Scripting.Dictionary")
    Dim uniquePduSizes As Object: Set uniquePduSizes = CreateObject("Scripting.Dictionary")
    Dim hasMultiplePduSizes As Boolean: hasMultiplePduSizes = False
    Dim sortedPduList() As Long
    
    If Not pduTable Is Nothing And lenIdx > 0 Then
        pduArr = pduTable.DataBodyRange.Value
        
        ' Build local map lookup table from ADU2NumSubchansTable
        ' Col 1 = LEN/ADU, Col 3 = PDU Size (B), Col 4 = MCS
        For i = 1 To UBound(pduArr, 1)
            Dim aduLenKey As Long: aduLenKey = CLng(pduArr(i, 1))
            Dim pduSizeVal As Long: pduSizeVal = CLng(pduArr(i, 3))
            Dim mcsVal As Long: mcsVal = CLng(pduArr(i, 4))
            lenToPduMap(aduLenKey) = pduSizeVal
            lenToMcsMap(aduLenKey) = mcsVal
        Next i
        
        ' Discover unique PDU Sizes actually present in ExpResultsTable
        For i = 1 To UBound(dataArr, 1)
            Dim curLen As Long: curLen = CLng(dataArr(i, lenIdx))
            If lenToPduMap.Exists(curLen) Then
                Dim targetPdu As Long: targetPdu = lenToPduMap(curLen)
                uniquePduSizes(targetPdu) = True
            End If
        Next i
        
        ' Execute N rows if unique PDU configuration space is > 1
        If uniquePduSizes.count > 1 Then
            hasMultiplePduSizes = True
            Dim pduKey As Variant, loopIdx As Long: loopIdx = 0
            ReDim sortedPduList(1 To uniquePduSizes.count)
            
            For Each pduKey In uniquePduSizes.Keys
                loopIdx = loopIdx + 1
                sortedPduList(loopIdx) = CLng(pduKey)
            Next pduKey
            
            ' Chronological array sort to prevent un-ordered structural plots
            Dim x1 As Long, x2 As Long, tempSwap As Long
            For x1 = 1 To UBound(sortedPduList) - 1
                For x2 = x1 + 1 To UBound(sortedPduList)
                    If sortedPduList(x1) > sortedPduList(x2) Then
                        tempSwap = sortedPduList(x1)
                        sortedPduList(x1) = sortedPduList(sortedPduList(x2))
                        sortedPduList(x2) = tempSwap
                    End If
                Next x2
            Next x1
            
            ' Iterate over discovered PDU values and call matching sub-arrays
            For loopIdx = 1 To UBound(sortedPduList)
                Dim activePduFilter As Long: activePduFilter = sortedPduList(loopIdx)
                rowTopAnchor = rowTopAnchor + 280
                
                UpdateProgressBar 20 + Int((loopIdx / UBound(sortedPduList)) * 10), "Processing RX PDU Size Sub-Block (" & activePduFilter & " B)..."
                
                ' Slice the master data array isolating matching PDU payloads
                Dim pduFilteredData() As Variant
                Dim filteredCount As Long: filteredCount = 0
                ReDim pduFilteredData(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))
                
                ' Find corresponding MCS for title naming purposes
                Dim activeMcsTitleVal As Long: activeMcsTitleVal = 0
                
                For i = 1 To UBound(dataArr, 1)
                    Dim trackingLen As Long: trackingLen = CLng(dataArr(i, lenIdx))
                    If lenToPduMap.Exists(trackingLen) Then
                        If lenToPduMap(trackingLen) = activePduFilter Then
                            filteredCount = filteredCount + 1
                            activeMcsTitleVal = lenToMcsMap(trackingLen)
                            Dim colWalk As Long
                            For colWalk = 1 To UBound(dataArr, 2)
                                pduFilteredData(filteredCount, colWalk) = dataArr(i, colWalk)
                            Next colWalk
                        End If
                    End If
                Next i
                
                If filteredCount > 0 Then
                    ' Pack rows tightly into a matching variants structure block
                    Dim finalSubData() As Variant: ReDim finalSubData(1 To filteredCount, 1 To UBound(dataArr, 2))
                    For i = 1 To filteredCount
                        For colWalk = 1 To UBound(dataArr, 2)
                            finalSubData(i, colWalk) = pduFilteredData(i, colWalk)
                        Next colWalk
                    Next i
                    
                    ' Pass structural tracking down into baseline analysis block handler
                    RunAnalysisBlock wsHist, finalSubData, rxCols, sfnIdx, txidIdx, "RX MAC LATENCY (PDU Size = " & activePduFilter & ", MCS = " & activeMcsTitleVal & ")", _
                                     minRx, maxRX, stepRX, _
                                     False, rowTopAnchor, IIf(doIndividualPlots, vbYes, vbNo), False, stToVenMap, uniqueVendors
                End If
            Next loopIdx
        End If
    End If
    ' ===========================================================================================

    UpdateProgressBar 35, "Processing Row 2: TX MAC Latency..."
    ' --- ROW 2: TX MAC LATENCY (Stations) ---
    rowTopAnchor = rowTopAnchor + 280
    RunAnalysisBlock wsHist, dataArr, rxCols, txqIdx, txidIdx, "TX MAC LATENCY", _
                     minTX, maxTX, stepTX, _
                     True, rowTopAnchor, IIf(doIndividualPlots, vbYes, vbNo), False, stToVenMap, uniqueVendors, sfnIdx

    ' ===========================================================================================
    ' INJECTED N-ROWS: UNIQUE PDU SIZE (B) FILTERED SUB-ANALYSIS FOR TX MAC LATENCY
    ' ===========================================================================================
    If hasMultiplePduSizes Then
        For loopIdx = 1 To UBound(sortedPduList)
            activePduFilter = sortedPduList(loopIdx)
            rowTopAnchor = rowTopAnchor + 280
            
            UpdateProgressBar 35 + Int((loopIdx / UBound(sortedPduList)) * 15), "Processing TX PDU Size Sub-Block (" & activePduFilter & " B)..."
            
            filteredCount = 0
            ReDim pduFilteredData(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))
            activeMcsTitleVal = 0
            
            For i = 1 To UBound(dataArr, 1)
                trackingLen = CLng(dataArr(i, lenIdx))
                If lenToPduMap.Exists(trackingLen) Then
                    If lenToPduMap(trackingLen) = activePduFilter Then
                        filteredCount = filteredCount + 1
                        activeMcsTitleVal = lenToMcsMap(trackingLen)
                        For colWalk = 1 To UBound(dataArr, 2)
                            pduFilteredData(filteredCount, colWalk) = dataArr(i, colWalk)
                        Next colWalk
                    End If
                End If
            Next i
            
            If filteredCount > 0 Then
                ReDim finalSubData(1 To filteredCount, 1 To UBound(dataArr, 2))
                For i = 1 To filteredCount
                    For colWalk = 1 To UBound(dataArr, 2)
                        finalSubData(i, colWalk) = pduFilteredData(i, colWalk)
                    Next colWalk
                Next i
                
                RunAnalysisBlock wsHist, finalSubData, rxCols, txqIdx, txidIdx, "TX MAC LATENCY (PDU Size = " & activePduFilter & ", MCS = " & activeMcsTitleVal & ")", _
                                 minTX, maxTX, stepTX, _
                                 True, rowTopAnchor, IIf(doIndividualPlots, vbYes, vbNo), False, stToVenMap, uniqueVendors, sfnIdx
            End If
        Next loopIdx
    End If
    ' ===========================================================================================

    UpdateProgressBar 55, "Processing Row 3: Total Latency..."
    ' --- ROW 3: TOTAL LATENCY (Stations) ---
    rowTopAnchor = rowTopAnchor + 280
    RunAnalysisBlock wsHist, dataArr, rxCols, genIdx, txidIdx, "TOTAL LATENCY", _
                     minTot, maxTot, stepTot, _
                     False, rowTopAnchor, IIf(doIndividualPlots, vbYes, vbNo), True, stToVenMap, uniqueVendors

    ' --- ROW 4: AC MAC LATENCY ---
    If msgIdIdx > 0 And acMap.count > 0 Then
        UpdateProgressBar 75, "Processing Row 4: AC MAC Latency..."
        rowTopAnchor = rowTopAnchor + 280
        RunACAnalysisBlock wsHist, dataArr, rxCols, txqIdx, msgIdIdx, acMap, "MAC LATENCY", _
                           (minRx + minTX), (maxRX + maxTX), stepRX, rowTopAnchor, False, stToVenMap, uniqueVendors
    End If

    ' --- ROW 5: AC TOTAL LATENCY ---
    If msgIdIdx > 0 And acMap.count > 0 Then
        UpdateProgressBar 90, "Processing Row 5: AC Total Latency..."
        rowTopAnchor = rowTopAnchor + 280
        RunACAnalysisBlock wsHist, dataArr, rxCols, genIdx, msgIdIdx, acMap, "TOTAL LATENCY", _
                           minTot, maxTot, stepTot, rowTopAnchor, True, stToVenMap, uniqueVendors
    End If

    UpdateProgressBar 95, "Formatting Report Layout..."
    wsHist.Columns("A:H").AutoFit
    wsHist.Range("B:H").NumberFormat = "0.00"
    
    ' Disable Cover / Restore Screen Settings
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    
    ' Clear the status bar
    Application.StatusBar = False
    
    ' Record processing time into global logging pipeline if active
    Dim tElapsed As Double: tElapsed = MicroTimer() - tStart
    If Not logTable Is Nothing Then
        logTable("GenerateLatencyAnalysis") = tElapsed
    End If
End Sub

' ===========================================================================================
' HELPERS: COVER & PROGRESS BAR
' ===========================================================================================
Private Sub CoverScreen(ByVal startPerformanceMode As Boolean)
    With Application
        If startPerformanceMode Then
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        Else
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

Private Sub UpdateProgressBar(ByVal percent As Integer, ByVal statusMsg As String)
    Dim barLength As Integer: barLength = 20
    Dim filledCount As Integer: filledCount = Round((percent / 100) * barLength)
    Dim emptyCount As Integer: emptyCount = barLength - filledCount
    
    Dim progressStr As String
    progressStr = "[" & String(filledCount, ChrW(&H2588)) & String(emptyCount, ChrW(&H2591)) & "]"
    
    Application.StatusBar = "Progress: " & progressStr & " " & percent & "% | " & statusMsg
    DoEvents
End Sub

' ===========================================================================================
' CORE PROCESSING LOGIC (Vendor Splicing Applied to Target Outputs)
' ===========================================================================================
Sub RunAnalysisBlock(ws As Worksheet, data As Variant, rxCols() As Long, baseIdx As Long, txidIdx As Long, title As String, _
                     bMin As Double, bMax As Double, bStep As Double, isTX As Boolean, topPos As Double, _
                     choice As VbMsgBoxResult, forceIntegerTicks As Boolean, stToVenMap As Object, uniqueVendors As Object, Optional sfnIdx As Long = 0)
    
    Dim nBins As Long: nBins = CLng((bMax - bMin) / bStep) + 1
    If nBins > 2000 Then nBins = 2000

    Dim n As Long, r As Long, countVal As Long, vIdx As Long
    Dim lats() As Double, startRow As Long, val As Double
    
    startRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 2
    If ws.Cells(1, 1).Value = "" Then startRow = 1
    ws.Cells(startRow, 1).Value = title
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow + 1, 1).Resize(1, 8).Value = Array("Station/Vendor", "MIN", "MAX", "MEAN", "Std. Dev.", "MODE", "95th %", "99th %")
    ws.Cells(startRow + 1, 1).Font.Bold = True

    Dim outRowOffset As Long: outRowOffset = 2
    Dim plotSlotIndex As Long: plotSlotIndex = 0

    ' 1. Separate Vendor Sub-Pool Blocks (Replaces Single OVERALL Matrix)
    Dim vKey As Variant
    For Each vKey In uniqueVendors.Keys
        ReDim lats(1 To UBound(data, 1) * UBound(rxCols)): countVal = 0
        For n = 1 To UBound(rxCols)
            If stToVenMap.Exists(CStr(n)) Then
                If stToVenMap(CStr(n)) = CStr(vKey) Then
                    For r = 1 To UBound(data, 1)
                        If (Not isTX) Or (isTX And data(r, txidIdx) = n) Then
                            If isTX Then val = data(r, sfnIdx) - data(r, baseIdx) Else val = data(r, rxCols(n)) - data(r, baseIdx)
                            If val >= 0 Then
                                countVal = countVal + 1
                                lats(countVal) = val
                            End If
                        End If
                    Next r
                End If
            End If
        Next n
        
        If countVal > 0 Then
            ReDim Preserve lats(1 To countVal)
            RenderData ws, lats, outRowOffset, nBins, bMin, bStep, "Vendor " & vKey, startRow, plotSlotIndex, bMin, bMax, topPos, "Vendor " & vKey & " " & title, forceIntegerTicks, True
            outRowOffset = outRowOffset + 1
            plotSlotIndex = plotSlotIndex + 1
        End If
    Next vKey

    ' 2. Conditional Station Breakdown Rows
    If choice = vbYes Then
        For n = 1 To UBound(rxCols)
            ReDim lats(1 To UBound(data, 1)): countVal = 0
            For r = 1 To UBound(data, 1)
                If (Not isTX) Or (isTX And data(r, txidIdx) = n) Then
                    If isTX Then val = data(r, sfnIdx) - data(r, baseIdx) Else val = data(r, rxCols(n)) - data(r, baseIdx)
                    If val >= 0 Then
                        countVal = countVal + 1: lats(countVal) = val
                    End If
                End If
            Next r
            If countVal > 0 Then
                ReDim Preserve lats(1 To countVal)
                RenderData ws, lats, outRowOffset, nBins, bMin, bStep, CStr(n), startRow, plotSlotIndex, bMin, bMax, topPos, "Station " & n & " " & title, forceIntegerTicks, False
                outRowOffset = outRowOffset + 1
                plotSlotIndex = plotSlotIndex + 1
            End If
        Next n
    End If
End Sub

Sub RunACAnalysisBlock(ws As Worksheet, data As Variant, rxCols() As Long, baseIdx As Long, msgIdIdx As Long, _
                       acMap As Object, title As String, bMin As Double, bMax As Double, bStep As Double, _
                       topPos As Double, forceIntegerTicks As Boolean, stToVenMap As Object, uniqueVendors As Object)
    
    Dim nBins As Long: nBins = CLng((bMax - bMin) / bStep) + 1
    If nBins > 2000 Then nBins = 2000

    Dim acVal As Integer, r As Long, n As Long, countVal As Long
    Dim lats() As Double, startRow As Long, val As Double, msgKey As String
    
    startRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 2
    ws.Cells(startRow, 1).Value = "AC " & title
    ws.Cells(startRow, 1).Font.Bold = True
    ws.Cells(startRow + 1, 1).Resize(1, 8).Value = Array("Station/Vendor", "MIN", "MAX", "MEAN", "Std. Dev.", "MODE", "95th %", "99th %")
    ws.Cells(startRow + 1, 1).Font.Bold = True

    Dim outRowOffset As Long: outRowOffset = 2
    Dim plotSlotIndex As Long: plotSlotIndex = 0

    ' 1. Separate Vendor Sub-Pool Blocks for AC Tracking Workspace
    Dim vKey As Variant
    For Each vKey In uniqueVendors.Keys
        For acVal = 0 To 3
            ReDim lats(1 To UBound(data, 1) * UBound(rxCols)): countVal = 0
            For r = 1 To UBound(data, 1)
                msgKey = Trim(CStr(data(r, msgIdIdx)))
                If acMap.Exists(msgKey) Then
                    If acMap(msgKey) = acVal Then
                        For n = 1 To UBound(rxCols)
                            If stToVenMap.Exists(CStr(n)) Then
                                If stToVenMap(CStr(n)) = CStr(vKey) Then
                                    If IsNumeric(data(r, rxCols(n))) And IsNumeric(data(r, baseIdx)) Then
                                        val = data(r, rxCols(n)) - data(r, baseIdx)
                                        If val >= 0 Then
                                            countVal = countVal + 1
                                            lats(countVal) = val
                                        End If
                                    End If
                                End If
                            End If
                        Next n
                    End If
                End If
            Next r
            
            If countVal > 0 Then
                ReDim Preserve lats(1 To countVal)
                RenderData ws, lats, outRowOffset, nBins, bMin, bStep, "Vendor " & vKey & " (AC " & acVal & ")", startRow, plotSlotIndex, bMin, bMax, topPos, "Vendor " & vKey & " AC " & acVal & " " & title, forceIntegerTicks, True
                outRowOffset = outRowOffset + 1
                plotSlotIndex = plotSlotIndex + 1
            End If
        Next acVal
    Next vKey
End Sub

Sub RenderData(ws As Worksheet, arr() As Double, rowOff As Long, nB As Long, bM As Double, bS As Double, _
               lbl As String, sRow As Long, slot As Long, xMin As Double, xMax As Double, topVal As Double, _
               chartTitleText As String, forceIntegerTicks As Boolean, isVendor As Boolean)
    
    QuickSort arr, LBound(arr), UBound(arr)
    Dim i As Long, sumV As Double, meanV As Double, sumSq As Double, stDevVal As Double
    Dim p95 As Double, p99 As Double, n As Long: n = UBound(arr)
    
    For i = 1 To n: sumV = sumV + arr(i): Next i
    meanV = sumV / n
    If n > 1 Then
        For i = 1 To n: sumSq = sumSq + (arr(i) - meanV) ^ 2: Next i
        stDevVal = Sqr(sumSq / (n - 1))
    Else: stDevVal = 0: End If
    
    p95 = arr(WorksheetFunction.Max(1, Int(n * 0.95)))
    p99 = arr(WorksheetFunction.Max(1, Int(n * 0.99)))
    
    ws.Cells(sRow + rowOff, 1).Resize(1, 8).Value = Array(lbl, arr(1), arr(n), meanV, stDevVal, GetMode(arr), p95, p99)

    If isVendor Then ws.Cells(sRow + rowOff, 1).Resize(1, 8).Font.Bold = True

    Dim xLabels() As Double, yFreq() As Double, yCDF() As Double
    ReDim xLabels(1 To nB): ReDim yFreq(1 To nB): ReDim yCDF(1 To nB)
    Dim bCounts() As Long: ReDim bCounts(1 To nB)
    For i = 1 To n
        Dim bIdx As Long
        If bS > 0 Then
            bIdx = Int((arr(i) - bM) / bS) + 1
            If bIdx >= 1 And bIdx <= nB Then bCounts(bIdx) = bCounts(bIdx) + 1
        End If
    Next i
    
    Dim curCum As Long: curCum = 0
    For i = 1 To nB: curCum = curCum + bCounts(i): xLabels(i) = bM + (i - 1) * bS: yFreq(i) = bCounts(i): yCDF(i) = curCum / n: Next i
    
    Dim cht As ChartObject: Set cht = ws.ChartObjects.Add(ws.Columns("J").Left + (slot * 380), topVal, 370, 260)
    With cht.Chart
        .HasTitle = True: .ChartTitle.Text = chartTitleText
        With .SeriesCollection.NewSeries
            .Name = "Frequency": .Values = yFreq: .XValues = xLabels: .ChartType = xlColumnClustered
        End With
        .ChartGroups(1).GapWidth = 50
        With .SeriesCollection.NewSeries
            .Name = "CDF": .Values = yCDF: .XValues = xLabels: .ChartType = xlXYScatterLinesNoMarkers: .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(0, 128, 0)
        End With
        With .SeriesCollection.NewSeries
            .Name = "95%": .XValues = Array(p95, p95): .Values = Array(0, 1): .ChartType = xlXYScatterLinesNoMarkers: .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0): .Format.Line.DashStyle = msoLineDash
        End With
        With .SeriesCollection.NewSeries
            .Name = "99%": .XValues = Array(p99, p99): .Values = Array(0, 1): .ChartType = xlXYScatterLinesNoMarkers: .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0): .Format.Line.DashStyle = msoLineDash
        End With
        
        With .Axes(xlCategory)
            .HasTitle = True: .AxisTitle.Text = "Time (ms)"
            .TickLabels.Orientation = 90
            
            If forceIntegerTicks Then
                .CategoryType = xlTimeScale
                .TickLabels.NumberFormat = "0"
                
                ' Calculate range
                Dim xMaxVal As Double
                xMaxVal = bM + (nB - 1) * bS
                
                ' Set clean, predictable integer step sizes
                If xMaxVal <= 50 Then
                    .MajorUnit = 5
                ElseIf xMaxVal <= 120 Then
                    .MajorUnit = 10
                Else
                    .MajorUnit = 20
                End If
            End If
        End With
        
        With .Axes(xlCategory, xlSecondary)
            .MinimumScale = bM - (bS / 2): .MaximumScale = (bM + (nB - 1) * bS) + (bS / 2)
            .TickLabelPosition = xlNone: .Format.Line.Visible = msoFalse
        End With
        With .Axes(xlValue): .HasTitle = True: .AxisTitle.Text = "Frequency": End With
        With .Axes(xlValue, xlSecondary): .HasTitle = True: .AxisTitle.Text = "CDF Probability": .MinimumScale = 0: .MaximumScale = 1: End With
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot As Double, tmpSwap As Double, tmpLow As Long, tmpHi As Long
    tmpLow = inLow
    tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    
    Do While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Loop
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

Function GetMode(arr() As Double) As Double
    Dim maxC As Long, curC As Long, mVal As Double, i As Long: maxC = 1: curC = 1: mVal = arr(LBound(arr))
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i) = arr(i - 1) Then curC = curC + 1 Else curC = 1
        If curC > maxC Then maxC = curC: mVal = arr(i)
    Next i
    GetMode = mVal
End Function

