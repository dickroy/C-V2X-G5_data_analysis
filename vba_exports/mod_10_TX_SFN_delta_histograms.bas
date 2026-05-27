Attribute VB_Name = "mod_10_TX_SFN_delta_histograms"
' ===========================================================================================
' Module: TX_SFN_delta_histograms
' Version: 4.1.0 - CONTIGUOUS MEMORY WITH AUTO-SORT RESTORATION
' Status: PERFORMANCE WRAPPERS, PROGRESS BAR, AND TIMING INTEGRATED
' ===========================================================================================

Option Explicit

' Native high-precision timing API
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
' WRAPPER FOR MANUAL EXECUTION
' ===========================================================================================
Sub Run_GenerateTX_SFN_delta_histograms()
    Dim startTime As Double: startTime = MicroTimer()
    Call GenerateTX_SFN_delta_histograms(Nothing)
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "TX SFN Delta Histogram Analysis Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT
' ===========================================================================================
Sub GenerateTX_SFN_delta_histograms(Optional ByRef logTable As Object = Nothing)
    Dim tStart As Double: tStart = MicroTimer()
    
    Dim wsSrc As Worksheet, wsHist As Worksheet
    Dim targetTable As ListObject
    Dim dataArr As Variant
    Dim txidIdx As Long, sfnIdx As Long, lenIdx As Long
    Dim rowTopAnchor As Double: rowTopAnchor = 20
    
    Const TARGET_SHEET As String = "TX_SFN delta Histograms"
    Const MAX_BIN As Long = 200
    
    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    On Error Resume Next
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    On Error GoTo 0
    If targetTable Is Nothing Then
        MsgBox "Error: Table 'ExpResultsTable' not found!", vbCritical
        Exit Sub
    End If

    ' Enable Performance Mode
    CoverScreen True
    UpdateProgressBar 5, "Initializing SFN structures..."

    ' Identify necessary column indices
    txidIdx = targetTable.ListColumns("TX_ID").Index
    sfnIdx = targetTable.ListColumns("TX_SFN_est").Index
    lenIdx = targetTable.ListColumns("LEN").Index

    ' ---------------------------------------------------------------------------------------
    ' STEP 1: SORT THE TABLE PHYSICALLY AND EXTRACT CONTIGUOUS DATA IN MEMORY
    ' ---------------------------------------------------------------------------------------
    UpdateProgressBar 10, "Optimizing table sort for extraction..."
    
    With targetTable.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=targetTable.ListColumns("TX_ID").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add2 key:=targetTable.ListColumns("TX_SFN_est").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    ' Pull contiguous data into memory array
    dataArr = targetTable.DataBodyRange.Value

    ' ---------------------------------------------------------------------------------------
    ' STEP 2: IMMEDIATELY RESTORE ORIGINAL SORT (TXQTIME) TO PREVENT PIPELINE BREAKS
    ' ---------------------------------------------------------------------------------------
    UpdateProgressBar 20, "Restoring table sort to TXQTIME..."
    
    Dim txqtimeCol As ListColumn
    On Error Resume Next
    Set txqtimeCol = targetTable.ListColumns("TXQTIME")
    On Error GoTo 0
    
    If Not txqtimeCol Is Nothing Then
        With targetTable.Sort
            .SortFields.Clear
            .SortFields.Add2 key:=txqtimeCol.DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    Else
        ' Fallback warning if column name varies slightly
        MsgBox "Warning: 'TXQTIME' column not found. Physical table order was not restored.", vbExclamation
    End If

    ' ---------------------------------------------------------------------------------------
    ' STEP 3: O(N) IN-MEMORY BOUNDARY MAPPING
    ' ---------------------------------------------------------------------------------------
    UpdateProgressBar 35, "Mapping transmitter boundaries in memory..."
    Dim txStartRows As Object: Set txStartRows = CreateObject("Scripting.Dictionary")
    Dim txEndRows As Object: Set txEndRows = CreateObject("Scripting.Dictionary")
    Dim lenDicts As Object: Set lenDicts = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, numRows As Long: numRows = UBound(dataArr, 1)
    Dim currentTX As String, lastTX As String: lastTX = ""
    
    For i = 1 To numRows
        currentTX = CStr(dataArr(i, txidIdx))
        
        If currentTX <> lastTX Then
            If lastTX <> "" Then txEndRows(lastTX) = i - 1
            txStartRows(currentTX) = i
            Set lenDicts(currentTX) = CreateObject("Scripting.Dictionary")
            lastTX = currentTX
        End If
        
        ' Track unique packet sizes (lengths)
        lenDicts(currentTX)(dataArr(i, lenIdx)) = True
    Next i
    If lastTX <> "" Then txEndRows(lastTX) = numRows

    ' Clean-delete the old sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(TARGET_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsHist = ThisWorkbook.Sheets.Add(After:=wsSrc)
    wsHist.Name = TARGET_SHEET

    ' Setup Report Headers
    Dim startRow As Long: startRow = 1
    wsHist.Cells(startRow, 1).Value = "TX SFN DELTA ANALYSIS"
    wsHist.Cells(startRow, 1).Font.Bold = True
    wsHist.Cells(startRow + 1, 1).Resize(1, 8).Value = Array("Station", "MIN", "MAX", "MEAN", "Std. Dev.", "MODE", "95th %", "99th %")
    wsHist.Cells(startRow + 1, 1).Font.Bold = True

    ' Process and Render Data (Direct Array Extraction - Instant O(1) lookups)
    Dim currentIDCount As Long: currentIDCount = 0
    Dim totalValid As Long: totalValid = txStartRows.count
    Dim key As Variant
    Dim sRow As Long, eRow As Long, segmentCount As Long
    Dim deltas() As Double, r As Long
    Dim lengths As Variant
    
    Dim lastPercent As Integer: lastPercent = 0
    Dim currentPercent As Integer
    
    For Each key In txStartRows.Keys
        sRow = txStartRows(key)
        eRow = txEndRows(key)
        segmentCount = eRow - sRow + 1
        
        If segmentCount > 1 Then
            currentIDCount = currentIDCount + 1
            
            ' Throttled Progress Bar Updates
            currentPercent = 45 + Round((currentIDCount / totalValid) * 45)
            If currentPercent <> lastPercent Then
                UpdateProgressBar currentPercent, "Processing TX_ID " & key & "..."
                lastPercent = currentPercent
            End If
            
            lengths = lenDicts(key).Keys
            QuickSortArray lengths, LBound(lengths), UBound(lengths)
            
            ' Directly populate deltas array from sorted memoryoffsets
            ReDim deltas(1 To segmentCount - 1)
            For r = 1 To segmentCount - 1
                deltas(r) = Abs(CLng(dataArr(sRow + r, sfnIdx)) - CLng(dataArr(sRow + r - 1, sfnIdx)))
            Next r
            
            ' Render statistics and chart
            RenderData wsHist, deltas, currentIDCount + 1, MAX_BIN, 0, 1, "Station " & key, startRow, _
                       currentIDCount - 1, 0, MAX_BIN, rowTopAnchor, _
                       "TX_ID " & key & " - SFN Delta Distribution - [" & Join(lengths, ", ") & "]"
                       
            rowTopAnchor = rowTopAnchor + 280
        End If
    Next key

QuickExit:
    UpdateProgressBar 95, "Formatting SFN Report Layout..."
    wsHist.Columns("A:H").AutoFit
    wsHist.Range("B:H").NumberFormat = "0.00"
    
    ' Disable Cover / Restore Screen Settings
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    Application.StatusBar = False
    
    Dim tElapsed As Double: tElapsed = MicroTimer() - tStart
    If Not logTable Is Nothing Then
        logTable("GenerateTX_SFN_delta_histograms") = tElapsed
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
' RENDER ENGINE
' ===========================================================================================
Sub RenderData(ws As Worksheet, ByRef arr() As Double, rowOff As Long, nB As Long, bM As Double, bS As Double, _
                lbl As String, sRow As Long, slot As Long, xMin As Double, xMax As Double, topVal As Double, _
                chartTitleText As String)
    
    ' 1. Sort the extracted segment array (highly optimized for size)
    QuickSortArray arr, LBound(arr), UBound(arr)
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

    ' 2. Bin distribution and Mode calculations
    Dim xLabels() As Double, yFreq() As Double, yCDF() As Double
    ReDim xLabels(1 To nB + 1): ReDim yFreq(1 To nB + 1): ReDim yCDF(1 To nB + 1)
    Dim bCounts() As Long: ReDim bCounts(1 To nB + 1)
    
    Dim maxBinCount As Long: maxBinCount = -1
    Dim modeVal As Double: modeVal = 0
    
    For i = 1 To n
        Dim bIdx As Long
        bIdx = Int((arr(i) - bM) / bS) + 1
        If bIdx >= 1 And bIdx <= nB Then
            bCounts(bIdx) = bCounts(bIdx) + 1
        ElseIf bIdx > nB Then
            bCounts(nB + 1) = bCounts(nB + 1) + 1
        End If
    Next i
    
    Dim curCum As Long: curCum = 0
    For i = 1 To nB
        curCum = curCum + bCounts(i)
        xLabels(i) = bM + (i - 1) * bS
        yFreq(i) = bCounts(i)
        yCDF(i) = curCum / n
        If bCounts(i) > maxBinCount Then
            maxBinCount = bCounts(i)
            modeVal = xLabels(i)
        End If
    Next i
    
    curCum = curCum + bCounts(nB + 1)
    xLabels(nB + 1) = nB + 1
    yFreq(nB + 1) = bCounts(nB + 1)
    yCDF(nB + 1) = curCum / n
    If bCounts(nB + 1) > maxBinCount Then
        modeVal = xLabels(nB + 1)
    End If
    
    ws.Cells(sRow + rowOff, 1).Resize(1, 8).Value = Array(lbl, arr(1), arr(n), meanV, stDevVal, modeVal, p95, p99)
    
    ' 3. Render Chart (Direct pixel placement)
    Const CHART_LEFT_PIXELS As Double = 520
    Dim cht As ChartObject: Set cht = ws.ChartObjects.Add(CHART_LEFT_PIXELS, topVal, 750, 260)
    
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
            .HasTitle = True: .AxisTitle.Text = "Delta (ms)"
            .TickLabels.Orientation = 90
            .TickLabelSpacing = 20
        End With
        
        With .Axes(xlCategory, xlSecondary)
            .MinimumScale = bM - (bS / 2)
            .MaximumScale = (bM + (nB) * bS) + (bS / 2)
            .TickLabelPosition = xlNone: .Format.Line.Visible = msoFalse
        End With
        
        With .Axes(xlValue): .HasTitle = True: .AxisTitle.Text = "Frequency": End With
        With .Axes(xlValue, xlSecondary): .HasTitle = True: .AxisTitle.Text = "CDF Probability": .MinimumScale = 0: .MaximumScale = 1: End With
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
End Sub

' ===========================================================================================
' SORTING UTILITIES
' ===========================================================================================
Private Sub QuickSortArray(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot As Variant, tmpSwap As Variant, tmpLow As Long, tmpHi As Long
    If inHi <= inLow Then Exit Sub
    tmpLow = inLow: tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi): tmpLow = tmpLow + 1: Wend
        While (pivot < vArray(tmpHi) And tmpHi > inLow): tmpHi = tmpHi - 1: Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow): vArray(tmpLow) = vArray(tmpHi): vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1: tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then QuickSortArray vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSortArray vArray, tmpLow, inHi
End Sub

