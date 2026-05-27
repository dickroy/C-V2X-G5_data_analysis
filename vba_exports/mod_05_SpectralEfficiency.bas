Attribute VB_Name = "mod_05_SpectralEfficiency"
Option Explicit

' ===========================================================================================
' HIGH-PRECISION TIMING API DECLARATIONS (Must be at the very top of the module)
' ===========================================================================================
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

' High-precision timer function (returns seconds with microsecond resolution)
Private Function MicroTimer() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency
    
    ' Call system APIs using strictly typed Currency variables
    If QueryPerformanceFrequency(cyFreq) <> 0 Then
        Call QueryPerformanceCounter(cyTicks)
        If cyFreq > 0 Then MicroTimer = cyTicks / cyFreq
    End If
End Function

' ===========================================================================================
' COVER ROUTINE: ENTRY POINT, PASSED TIMING METRICS, & PERFORMANCE MONITOR
' ===========================================================================================
Sub Run_GenerateSpectralEfficiencyAnalysis()
    Dim startTime As Double: startTime = MicroTimer()
    
    ' Clear and set initial status bar message
    Application.StatusBar = "Initializing C-V2X Spectral Efficiency Analysis..."
    
    ' Create a Scripting Dictionary to pass back performance telemetry from the core engine
    Dim runTelemetry As Object
    Set runTelemetry = CreateObject("Scripting.Dictionary")
    
    ' Call the core calculation engine
    Call GenerateSpectralEfficiencyAnalysis(runTelemetry)
    
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    
    ' Restore system status bar
    Application.StatusBar = False
    
    Dim coreTime As Double
    If runTelemetry.Exists("Core_Execution_Seconds") Then
        coreTime = runTelemetry("Core_Execution_Seconds")
    End If
    
    ' Display execution summary
    MsgBox "C-V2X Spectral Efficiency Analysis Complete." & vbCrLf & _
           "Core Calculation Engine: " & Format(coreTime, "0.000") & " seconds" & vbCrLf & _
           "Total Process Runtime (inc. Render): " & Format(totalRunTime, "0.000") & " seconds", _
           vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' CORE ENGINE: C-V2X PHYSICAL RESOURCE SPACE SPECTRAL EFFICIENCY
' ===========================================================================================
Sub GenerateSpectralEfficiencyAnalysis(Optional ByRef logTable As Object = Nothing)
    Dim tStartExec As Double: tStartExec = MicroTimer()
    
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim targetTable As ListObject, dgpsTable As ListObject
    Dim dataArr As Variant, dgpsArr As Variant
    Dim i As Long, rowCount As Long, dgpsRowCount As Long
    Dim colSFN As Long, colLEN As Long, colRXCOUNT As Long
    
    Dim tStart As Double, tStop As Double, tWin As Double, tStep As Double
    Dim numRx As Integer, TBW As Double
    Dim maxB_MCS7_Bytes As Double, maxB_MCS11_Bytes As Double
    Dim maxCapQPSK As Double, maxCap16QAM As Double
    Dim drawDeGps As Integer
    
    ' 1. CLEANUP & SETUP
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Spectral Efficiency").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    
    ' Reference deGPSTable on the "Test Description" sheet safely
    On Error Resume Next
    Set dgpsTable = ThisWorkbook.Sheets("Test Description").ListObjects("deGPSTable")
    On Error GoTo 0
    
    Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
    wsDest.Name = "Spectral Efficiency"
    
    ' Disable screen updating/calculations for raw speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 2. FETCH PARAMETERS & CONFIG (From Spreadsheet Named Ranges)
    On Error Resume Next
    tStart = Range("T_start").Value
    tStop = Range("T_stop").Value
    tWin = Range("T_win_size").Value
    tStep = Range("T_step").Value
    numRx = Range("Num_Rx_Stations").Value
    TBW = Range("TBWperSF").Value
    maxB_MCS7_Bytes = Range("MAXB_MCS7").Value
    maxB_MCS11_Bytes = Range("MAXB_MCS11").Value
    drawDeGps = Range("deGPS").Value ' Master toggle (1 = Draw, 0 = Skip)
    On Error GoTo 0
    
    ' --- CRITICAL UNITS AUTO-CORRECTION TO MILLISECONDS ---
    If tStart > 0 And tStart < 100 Then tStart = tStart * 1000
    If tStop > 0 And tStop < 100 Then tStop = tStop * 1000
    If tWin > 0 And tWin < 100 Then tWin = tWin * 1000
    If tStep > 0 And tStep < 100 Then tStep = tStep * 1000
    
    ' Fallbacks if named ranges are missing or empty
    If tWin = 0 Then tWin = 100
    If tStep = 0 Then tStep = 50
    If tStop = 0 Then tStop = 1000
    If numRx = 0 Then numRx = 6
    If TBW = 0 Then TBW = 20000
    If maxB_MCS7_Bytes = 0 Then maxB_MCS7_Bytes = 1479
    If maxB_MCS11_Bytes = 0 Then maxB_MCS11_Bytes = 2124
    
    ' Convert limits (bps/Hz)
    maxCapQPSK = ((maxB_MCS7_Bytes * 8) * (numRx - 1)) / TBW
    maxCap16QAM = ((maxB_MCS11_Bytes * 8) * (numRx - 1)) / TBW
    
    Application.StatusBar = "Sorting source table data by TX_SFN_est..."
    
    ' 3. SORT TABLE BY TX_SFN_EST TO GUARANTEE CONTIGUOUS O(1) SFN PARSING
    With targetTable.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=targetTable.ListColumns("TX_SFN_est").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    dataArr = targetTable.DataBodyRange.Value
    colSFN = targetTable.ListColumns("TX_SFN_est").Index
    colLEN = targetTable.ListColumns("LEN").Index
    colRXCOUNT = targetTable.ListColumns("RX_COUNT").Index
    rowCount = UBound(dataArr, 1)
    
    ' 4. PRE-INDEX CONTIGUOUS SFN ROW BOUNDARIES FOR INSTANT LOOKUPS
    Application.StatusBar = "Indexing subframes..."
    Dim sfnStartRows As Object: Set sfnStartRows = CreateObject("Scripting.Dictionary")
    Dim sfnEndRows As Object: Set sfnEndRows = CreateObject("Scripting.Dictionary")
    
    Dim currentSFN As Double, lastSFN As Double: lastSFN = -999999
    For i = 1 To rowCount
        currentSFN = CDbl(dataArr(i, colSFN))
        If currentSFN <> lastSFN Then
            If lastSFN <> -999999 Then sfnEndRows(lastSFN) = i - 1
            sfnStartRows(currentSFN) = i
            lastSFN = currentSFN
        End If
    Next i
    If lastSFN <> -999999 Then sfnEndRows(lastSFN) = rowCount
    
    ' Extract keys to a strongly-typed Double array for lightning-fast Binary Search
    Dim sfnKeys() As Double
    Dim totalSFNKeys As Long: totalSFNKeys = sfnStartRows.count
    ReDim sfnKeys(0 To totalSFNKeys - 1)
    
    Dim keyVar As Variant, kIdx As Long: kIdx = 0
    For Each keyVar In sfnStartRows.Keys
        sfnKeys(kIdx) = CDbl(keyVar)
        kIdx = kIdx + 1
    Next keyVar
    
    ' ---------------------------------------------------------------------------------------
    ' 4b. PARSE deGPS TABLE TIME VALUES DIRECTLY
    ' ---------------------------------------------------------------------------------------
    Dim dGpsOutRow As Long: dGpsOutRow = 2
    Dim dGpsCol As Integer: dGpsCol = 16 ' Column P for dGPS event tables
    
    wsDest.Cells(1, dGpsCol).Resize(1, 2).Value = Array("deGPS Event Time (s)", "Y-Anchor")
    wsDest.Cells(1, dGpsCol).Resize(1, 2).Font.Bold = True
    
    If drawDeGps = 1 And Not dgpsTable Is Nothing Then
        dgpsArr = dgpsTable.DataBodyRange.Value
        dgpsRowCount = UBound(dgpsArr, 1)
        
        For i = 1 To dgpsRowCount
            Dim eventTimeSec As Double
            eventTimeSec = CDbl(dgpsArr(i, 3))
            
            If (eventTimeSec * 1000) >= tStart And (eventTimeSec * 1000) <= tStop Then
                wsDest.Cells(dGpsOutRow, dGpsCol).Value = eventTimeSec
                wsDest.Cells(dGpsOutRow, dGpsCol + 1).Value = 0
                dGpsOutRow = dGpsOutRow + 1
            End If
        Next i
    End If
    Dim lastDgpsRow As Long: lastDgpsRow = dGpsOutRow - 1
    
    ' ---------------------------------------------------------------------------------------
    ' 5. SLIDING WINDOW PROCESSOR WITH BINARY SEARCH (O(log K) Key Matching)
    ' ---------------------------------------------------------------------------------------
    Dim outRow As Long: outRow = 2
    Dim dataStartCol As Integer: dataStartCol = 10 ' Column J
    
    ' Size and prepare output array in memory to avoid writing to cells row-by-row
    Dim maxOutRows As Long: maxOutRows = Int((tStop - tStart - tWin) / tStep) + 2
    Dim outBuffer() As Variant
    ReDim outBuffer(1 To maxOutRows, 1 To 5)
    
    outBuffer(1, 1) = "Mid Time (s)"
    outBuffer(1, 2) = "Measured (bps/Hz)"
    outBuffer(1, 3) = "Max QPSK (Blue Line)"
    outBuffer(1, 4) = "Max 16-QAM (Red Line)"
    outBuffer(1, 5) = "Subframes Present"
    
    Dim winStart As Double, winEnd As Double
    Dim winTotalBits As Double, midTimeSec As Double, specEff As Double
    
    winStart = tStart
    Do While (winStart + tWin) <= tStop
        winEnd = winStart + tWin
        winTotalBits = 0
        
        ' --- BINARY SEARCH FIND START SFN INDEX ---
        Dim low As Long: low = 0
        Dim high As Long: high = totalSFNKeys - 1
        Dim midIdx As Long
        Dim startMatchIdx As Long: startMatchIdx = -1
        
        ' Find the first SFN index >= winStart
        Do While low <= high
            midIdx = (low + high) \ 2
            If sfnKeys(midIdx) >= winStart Then
                startMatchIdx = midIdx
                high = midIdx - 1 ' Keep searching left for exact first boundary
            Else
                low = midIdx + 1
            End If
        Loop
        
        ' --- SUM BITS IN WINDOW ---
        Dim nSubfrs As Long: nSubfrs = 0
        If startMatchIdx <> -1 Then
            Dim sfnIdx As Long: sfnIdx = startMatchIdx
            Do While sfnIdx < totalSFNKeys
                Dim activeSFN As Double: activeSFN = sfnKeys(sfnIdx)
                If activeSFN < winEnd Then
                    ' Sum this SFN's pre-indexed row ranges
                    Dim sRow As Long: sRow = sfnStartRows(activeSFN)
                    Dim eRow As Long: eRow = sfnEndRows(activeSFN)
                    For i = sRow To eRow
                        winTotalBits = winTotalBits + (dataArr(i, colRXCOUNT) * dataArr(i, colLEN) * 8)
                    Next i
                    nSubfrs = nSubfrs + 1
                    sfnIdx = sfnIdx + 1
                Else
                    Exit Do ' Out of window bounds, skip remaining SFNs
                End If
            Loop
        End If
        
        midTimeSec = (winStart + (tWin / 2)) / 1000
        specEff = IIf(nSubfrs > 0, winTotalBits / (TBW * nSubfrs), 0)
        
        ' Save to buffer array
        outBuffer(outRow, 1) = midTimeSec
        outBuffer(outRow, 2) = specEff
        outBuffer(outRow, 3) = maxCapQPSK
        outBuffer(outRow, 4) = maxCap16QAM
        outBuffer(outRow, 5) = nSubfrs
        
        outRow = outRow + 1
        winStart = winStart + tStep
    Loop
    
    ' Dump calculations buffer array to sheet in a single write operation
    Dim lastRow As Long: lastRow = outRow - 1
    wsDest.Cells(1, dataStartCol).Resize(lastRow, 5).Value = outBuffer
    
    ' 6. SAFE FORMATTING & RESIZING
    If lastRow >= 2 Then
        wsDest.Cells(2, dataStartCol).Resize(lastRow - 1, 1).NumberFormat = "0.000"
        wsDest.Cells(2, dataStartCol + 1).Resize(lastRow - 1, 3).NumberFormat = "0.00"
        wsDest.Columns(dataStartCol).Resize(, 5).AutoFit
        If drawDeGps = 1 And lastDgpsRow >= 2 Then
            wsDest.Cells(2, dGpsCol).Resize(lastDgpsRow - 1, 1).NumberFormat = "0.000"
            wsDest.Columns(dGpsCol).Resize(, 2).AutoFit
        End If
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "No SFN data matched. Check your T_start / T_stop configuration.", vbExclamation, "No Data Found"
        Exit Sub
    End If
    
    ' 7. PLOT GENERATION
    Application.StatusBar = "Generating plots..."
    Dim specChart As ChartObject
    Set specChart = wsDest.ChartObjects.Add(Left:=10, Top:=10, Width:=750, Height:=450)
    
    With specChart.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True
        .ChartTitle.Text = "C-V2X Spectral Efficiency vs. Time"
        
        ' Series 1: Measured
        With .SeriesCollection.NewSeries
            .Name = "Measured"
            .XValues = wsDest.Range(wsDest.Cells(2, dataStartCol), wsDest.Cells(lastRow, dataStartCol))
            .Values = wsDest.Range(wsDest.Cells(2, dataStartCol + 1), wsDest.Cells(lastRow, dataStartCol + 1))
            .Format.Line.ForeColor.RGB = RGB(180, 180, 180)
            .Format.Line.Weight = 1.25
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6
        End With
        
        ' Series 2: Max Capacity - QPSK (BLUE LINE)
        With .SeriesCollection.NewSeries
            .Name = "Max Capacity (QPSK)"
            .XValues = wsDest.Range(wsDest.Cells(2, dataStartCol), wsDest.Cells(lastRow, dataStartCol))
            .Values = wsDest.Range(wsDest.Cells(2, dataStartCol + 2), wsDest.Cells(lastRow, dataStartCol + 2))
            .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        ' Series 3: Max Capacity - 16-QAM (RED LINE)
        With .SeriesCollection.NewSeries
            .Name = "Max Capacity (16-QAM)"
            .XValues = wsDest.Range(wsDest.Cells(2, dataStartCol), wsDest.Cells(lastRow, dataStartCol))
            .Values = wsDest.Range(wsDest.Cells(2, dataStartCol + 3), wsDest.Cells(lastRow, dataStartCol + 3))
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        ' ---------------------------------------------------------------------------------------
        ' 7b. deGPS STATE INTERRUPT EVENT SERIES & VERTICAL LINES (FIXED ERROR BAR METHOD)
        ' ---------------------------------------------------------------------------------------
        If drawDeGps = 1 And lastDgpsRow >= 2 Then
            With .SeriesCollection.NewSeries
                .Name = "deGPS Interrupt State"
                .XValues = wsDest.Range(wsDest.Cells(2, dGpsCol), wsDest.Cells(lastDgpsRow, dGpsCol))
                .Values = wsDest.Range(wsDest.Cells(2, dGpsCol + 1), wsDest.Cells(lastDgpsRow, dGpsCol + 1))
                .MarkerStyle = xlMarkerStyleNone
                .Format.Line.Visible = msoFalse
                
                ' Pass Literal Value '2' for xlErrorBarIncludePlus to evade missing definitions
                .ErrorBar Direction:=xlY, Include:=2, _
                          Type:=xlFixedValue, Amount:=maxCap16QAM * 1.15
                
                With .ErrorBars
                    .EndStyle = xlNoCap
                    With .Format.Line
                        .ForeColor.RGB = RGB(150, 0, 0)
                        .DashStyle = msoLineDash
                        .Weight = 1.25
                    End With
                End With
            End With
        End If

        ' Axes configurations
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Time (seconds)"
            .TickLabels.NumberFormat = "0.0"
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Spectral Efficiency (bps/Hz)"
            .MinimumScale = 0
            If maxCap16QAM > 0 Then .MaximumScale = WorksheetFunction.RoundUp(maxCap16QAM * 1.15, 1)
        End With
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' ---------------------------------------------------------------------------------------
        ' 7c. DYNAMIC TEXT LABELS: SYNC LOSS(i) UNDER H-AXIS
        ' ---------------------------------------------------------------------------------------
        DoEvents ' Yield thread to ensure positioning maps render accurately
        
        If drawDeGps = 1 And lastDgpsRow >= 2 Then
            Dim plotLeft As Double, plotWidth As Double, plotTop As Double, plotHeight As Double
            Dim xMin As Double, xMax As Double, xVal As Double, xPos As Double
            Dim syncLabel As Shape
            Dim eventIdx As Long
            
            plotLeft = .PlotArea.InsideLeft
            plotWidth = .PlotArea.InsideWidth
            plotTop = .PlotArea.InsideTop
            plotHeight = .PlotArea.InsideHeight
            
            xMin = .Axes(xlCategory).MinimumScale
            xMax = .Axes(xlCategory).MaximumScale
            
            For eventIdx = 2 To lastDgpsRow
                xVal = wsDest.Cells(eventIdx, dGpsCol).Value
                
                If xMax <> xMin And xVal >= xMin And xVal <= xMax Then
                    xPos = plotLeft + ((xVal - xMin) / (xMax - xMin)) * plotWidth
                    
                    Set syncLabel = wsDest.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                        specChart.Left + xPos - 35, _
                        specChart.Top + plotTop + plotHeight + 18, _
                        70, 18)
                    
                    With syncLabel
                        .TextFrame.Characters.Text = "SYNC LOSS(" & (eventIdx - 1) & ")"
                        
                        ' Fix for 424 Error: Use TextFrame2 to format text characteristics safely
                        With .TextFrame2.TextRange.Font
                            .Size = 8
                            .Bold = msoTrue
                            .Fill.ForeColor.RGB = RGB(150, 0, 0)
                        End With
                        
                        .TextFrame.HorizontalAlignment = xlHAlignCenter
                        .Line.Visible = msoFalse
                        .Fill.Transparency = 1#
                    End With
                End If
            Next eventIdx
        End If

        ' ---------------------------------------------------------------------------------------
        ' 8. HEATMAP COLORIZATION ENGINE (O(N) Render Optimization)
        ' ---------------------------------------------------------------------------------------
        Dim pts As Points: Set pts = .SeriesCollection("Measured").Points
        Dim ptIdx As Long
        Dim curEff As Double
        Dim ratio As Double
        Dim r As Integer, g As Integer, b As Integer
        
        For ptIdx = 1 To pts.count
            curEff = wsDest.Cells(ptIdx + 1, dataStartCol + 1).Value
            
            If curEff >= maxCapQPSK Then
                r = 0
                g = 220
                b = 0
            Else
                ratio = curEff / maxCapQPSK
                If ratio >= 0.66 Then
                    Dim localRatio1 As Double: localRatio1 = (ratio - 0.66) / 0.34
                    r = Int(255 * (1 - localRatio1))
                    g = 220
                    b = 0
                ElseIf ratio >= 0.33 Then
                    Dim localRatio2 As Double: localRatio2 = (ratio - 0.33) / 0.33
                    r = 255
                    g = Int(128 + (92 * localRatio2))
                    b = 0
                Else
                    Dim localRatio3 As Double: localRatio3 = ratio / 0.33
                    r = 255
                    g = Int(128 * localRatio3)
                    b = 0
                End If
            End If
            
            With pts(ptIdx)
                .MarkerBackgroundColor = RGB(r, g, b)
                .MarkerForegroundColor = RGB(r, g, b)
            End With
        Next ptIdx
    End With
    
    ' Re-enable screen updating/calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Write run telemetry metrics back to Cover Routine dictionary
    If Not logTable Is Nothing Then
        Dim tElapsedTotal As Double: tElapsedTotal = MicroTimer() - tStartExec
        logTable("Core_Execution_Seconds") = tElapsedTotal
    End If
End Sub

