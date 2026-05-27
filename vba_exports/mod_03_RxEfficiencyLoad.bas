Attribute VB_Name = "mod_03_RxEfficiencyLoad"
' Module: C_V2X_Efficiency_Analysis_2026
' Version: 1.3.1-INTEGRATED - TOTAL INTEGER ALIGNMENT FOR RX EFF TABLE
' ===========================================================================================

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
Sub Run_GenerateLoadRxEfficiencyAnalysis()
    Dim startTime As Double: startTime = MicroTimer()
    
    ' Call core macro with pipeline parameters set to manual state
    Call GenerateLoadRxEfficiencyAnalysis(Nothing)
    
    ' Output execution timing only on manual runs
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "Rx Load & Efficiency Analysis Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT (Safe for automated pipelines calling GenerateLoadRxEfficiencyAnalysis(Nothing))
' ===========================================================================================
Sub GenerateLoadRxEfficiencyAnalysis(Optional ByRef logTable As Object = Nothing)
    Dim tStart As Double: tStart = MicroTimer() ' Start Timing
    
    Dim wsSrc As Worksheet, wsDest As Worksheet, wsApp As Worksheet, wsPdu As Worksheet, wsDesc As Worksheet
    Dim targetTable As ListObject, pduTable As ListObject, appTable As ListObject, deGPSTable As ListObject
    Dim dataArr As Variant, pduArr As Variant, resultsArr As Variant, appArr As Variant, deGPSData As Variant
    Dim i As Long, j As Long, k As Long
    Dim tStartVal As Double, tStopVal As Double, tWin As Double, tStep As Double
    Dim numRxMax As Long, nSchPerSubframe As Double, minNschPerTx As Double
    Dim yMaxLimit As Double, dataMax As Double, finalYMax As Double
    Dim numSchSec As Double
    
    Dim colSFN As Long, colRxC As Long, colTXperSFN As Long, colLEN As Long
    Dim sfnGroupRx As Object: Set sfnGroupRx = CreateObject("Scripting.Dictionary")
    Dim nSchMap As Object: Set nSchMap = CreateObject("Scripting.Dictionary")
    Dim chtObjTime As ChartObject, chtObjLoad As ChartObject, chtObjTX As ChartObject

    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    Set wsApp = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
    Set wsPdu = ThisWorkbook.Sheets("PDU Size Table")
    
    ' Set the sheet containing the deGPS table
    On Error Resume Next
    Set wsDesc = ThisWorkbook.Sheets("Test Description")
    On Error GoTo 0
    
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    Set pduTable = wsPdu.ListObjects("ADU2NumSubchansTable")
    Set appTable = wsApp.ListObjects("AppParams")
    
    ' --- Target deGPSTable on the 'Test Description' sheet ---
    On Error Resume Next
    If Not wsDesc Is Nothing Then
        Set deGPSTable = wsDesc.ListObjects("deGPSTable")
    End If
    
    ' Fallbacks
    If deGPSTable Is Nothing Then Set deGPSTable = wsSrc.ListObjects("deGPSTable")
    If deGPSTable Is Nothing Then Set deGPSTable = wsApp.ListObjects("deGPSTable")
    On Error GoTo 0
    
    ' Enable Performance Settings & Begin UI Feedback
    CoverScreen True
    UpdateProgressBar 5, "Initializing Data Structures..."
    
    ' --- Retrieve Parameters ---
    tStartVal = wsApp.Range("T_start").Value
    tStopVal = wsApp.Range("T_stop").Value
    tWin = wsApp.Range("T_win_size").Value
    tStep = wsApp.Range("T_step").Value
    numRxMax = [Num_Rx_Stations]
    nSchPerSubframe = wsApp.Range("Nsch_per_subfr").Value
    minNschPerTx = wsApp.Range("Min_Nsch_per_tx").Value
    numSchSec = wsApp.Range("Num_SCH_sec").Value
    
    yMaxLimit = nSchPerSubframe / minNschPerTx
    
    dataArr = targetTable.DataBodyRange.Value
    pduArr = pduTable.DataBodyRange.Value
    colSFN = targetTable.ListColumns("TX_SFN_est").Index
    colRxC = targetTable.ListColumns("RX_COUNT").Index
    colTXperSFN = targetTable.ListColumns("TXperSFN").Index
    colLEN = targetTable.ListColumns("LEN").Index
    
    For i = 1 To UBound(pduArr, 1)
        nSchMap(CLng(pduArr(i, 1))) = CDbl(pduArr(i, 2))
    Next i
    
    For i = 1 To UBound(dataArr, 1)
        sfnGroupRx(dataArr(i, colSFN)) = sfnGroupRx(dataArr(i, colSFN)) + CDbl(dataArr(i, colRxC))
    Next i
    
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Load_Rx Efficiency Analysis")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "Load_Rx Efficiency Analysis"
    End If
    wsDest.ChartObjects.Delete
    On Error GoTo 0
    wsDest.Cells.Clear

    ' --- Calculate Windowed Results ---
    UpdateProgressBar 20, "Calculating windowed efficiency metrics..."
    Dim winCount As Long: winCount = Int((tStopVal - tStartVal) / tStep) + 1
    ReDim resultsArr(1 To winCount, 1 To 5)
    Dim ptrStart As Long: ptrStart = 1: Dim ptrEnd As Long: ptrEnd = 1
    
    dataMax = 0
    For k = 1 To winCount
        Dim cL As Double: cL = tStartVal + (k - 1) * tStep
        Dim cU As Double: cU = cL + tWin
        If cU > tStopVal Then Exit For
        
        Dim sumRxEff As Double: sumRxEff = 0
        Dim sumCVRxEff As Double: sumCVRxEff = 0
        Dim sumNsch As Double: sumNsch = 0
        Dim sumTXperSFN As Double: sumTXperSFN = 0
        Dim cnt As Long: cnt = 0
        
        Do While ptrStart <= UBound(dataArr, 1)
            If dataArr(ptrStart, colSFN) >= cL Then Exit Do
            ptrStart = ptrStart + 1
        Loop
        If ptrEnd < ptrStart Then ptrEnd = ptrStart
        Do While ptrEnd <= UBound(dataArr, 1)
            If dataArr(ptrEnd, colSFN) > cU Then Exit Do
            ptrEnd = ptrEnd + 1
        Loop
        
        If ptrEnd > ptrStart Then
            For i = ptrStart To ptrEnd - 1
                sumRxEff = sumRxEff + (CDbl(dataArr(i, colRxC)) / (numRxMax - 1))
                Dim txN As Long: txN = dataArr(i, colTXperSFN)
                sumTXperSFN = sumTXperSFN + txN
                If (numRxMax - txN) * txN > 0 Then
                    sumCVRxEff = sumCVRxEff + (sfnGroupRx(dataArr(i, colSFN)) / ((numRxMax - txN) * txN))
                End If
                cnt = cnt + 1
                If nSchMap.Exists(CLng(dataArr(i, colLEN))) Then
                    sumNsch = sumNsch + nSchMap(CLng(dataArr(i, colLEN)))
                End If
            Next i
        End If
        
        resultsArr(k, 1) = cU / 1000
        If cnt > 0 Then
            resultsArr(k, 2) = (sumRxEff / cnt)
            resultsArr(k, 3) = (sumCVRxEff / cnt)
            resultsArr(k, 5) = (sumTXperSFN / cnt)
            If resultsArr(k, 5) > dataMax Then dataMax = resultsArr(k, 5)
        End If
        resultsArr(k, 4) = (sumNsch / (tWin * nSchPerSubframe))
    Next k
    
    finalYMax = IIf(dataMax > yMaxLimit, dataMax, yMaxLimit)
    wsDest.Range("AI1:AM1").Value = Array("Time (s)", "Rx Efficiency", "C-V2X Rx Efficiency", "Load", "Filt TXperSFN")
    wsDest.Range("AI2").Resize(winCount, 5).Value = resultsArr
    
    wsDest.Range("AI2:AI" & winCount + 1).NumberFormat = "0.0"
    wsDest.Range("AJ2:AL" & winCount + 1).NumberFormat = "0.00%"
    wsDest.Range("AM2:AM" & winCount + 1).NumberFormat = "0.00"

    ' --- DEMAND CURVE LOGIC ---
    UpdateProgressBar 45, "Processing Application Demand metrics..."
    Dim lc As ListColumn, cleanName As String
    Dim colSID As Long, colADU As Long, colTTI As Long, colMID_App As Long
    Dim hasValidApp As Boolean: hasValidApp = False
    
    For Each lc In appTable.ListColumns
        cleanName = UCase(Trim(lc.Name))
        Select Case cleanName
            Case "STATION ID", "STATION_ID": colSID = lc.Index
            Case "ADU SIZE (B)", "ADU_SIZE": colADU = lc.Index
            Case "TTI(MS)", "TTI": colTTI = lc.Index
            Case "APP_ID", "MSG_ID", "IVI_ID": colMID_App = lc.Index
        End Select
    Next lc

    If Not (appTable.DataBodyRange Is Nothing) Then
        appArr = appTable.DataBodyRange.Value
        If colMID_App > 0 And colSID > 0 And colADU > 0 And colTTI > 0 Then hasValidApp = True
    End If

    Dim eCount As Long: eCount = 0
    Dim stepRes() As Variant

    If hasValidApp Then
        Dim colResMID As Long, colResTime As Long
        colResMID = targetTable.ListColumns("App_ID").Index
        colResTime = targetTable.ListColumns("TXQTIME").Index
        
        Dim rawEvents() As Variant: ReDim rawEvents(1 To UBound(appArr, 1) * 2, 1 To 2)
        
        ' App Summary Table (Rows 1:N, Cols A:D)
        wsDest.Range("A1:D1").Value = Array("App_ID", "T_start (s)", "T_stop (s)", "Demand")
        wsDest.Range("A1:D1").HorizontalAlignment = xlCenter
        
        For i = 1 To UBound(appArr, 1)
            Dim curApp As Variant: curApp = appArr(i, colMID_App)
            Dim firstT As Double: firstT = 0
            Dim lastT As Double: lastT = 0
            
            For k = 1 To UBound(dataArr, 1)
                If dataArr(k, colResMID) = curApp Then
                    If firstT = 0 Then firstT = dataArr(k, colResTime)
                    lastT = dataArr(k, colResTime)
                End If
            Next k
            
            Dim nSchs As Double: nSchs = 0
            If nSchMap.Exists(CLng(appArr(i, colADU))) Then nSchs = nSchMap(CLng(appArr(i, colADU)))
            
            Dim curDemand As Double: curDemand = (1000 * nSchs / appArr(i, colTTI)) / numSchSec
            
            wsDest.Cells(i + 1, 1).Value = curApp
            wsDest.Cells(i + 1, 2).Value = firstT / 1000
            wsDest.Cells(i + 1, 3).Value = lastT / 1000
            wsDest.Cells(i + 1, 4).Value = curDemand
            
            eCount = eCount + 1
            rawEvents(eCount, 1) = firstT / 1000: rawEvents(eCount, 2) = curDemand
            eCount = eCount + 1
            rawEvents(eCount, 1) = lastT / 1000: rawEvents(eCount, 2) = -curDemand
        Next i
        
        ' Apply explicit column formats to the populated App Summary Table range
        wsDest.Range("B2:C" & UBound(appArr, 1) + 1).NumberFormat = "0.0"
        wsDest.Range("D2:D" & UBound(appArr, 1) + 1).NumberFormat = "0.00%"
        
        ' Sort Events chronologically
        For i = 1 To eCount - 1
            For j = i + 1 To eCount
                If rawEvents(i, 1) > rawEvents(j, 1) Then
                    Dim tempT As Double: tempT = rawEvents(i, 1)
                    Dim tempD As Double: tempD = rawEvents(i, 2)
                    rawEvents(i, 1) = rawEvents(j, 1): rawEvents(i, 2) = rawEvents(j, 2)
                    rawEvents(j, 1) = tempT: rawEvents(j, 2) = tempD
                End If
            Next j
        Next i
        
        ' Generate Staircase (AV:AW)
        ReDim stepRes(1 To (eCount * 2), 1 To 2)
        Dim curSum As Double: curSum = 0: Dim pPtr As Long: pPtr = 1
        For i = 1 To eCount
            stepRes(pPtr, 1) = rawEvents(i, 1): stepRes(pPtr, 2) = curSum: pPtr = pPtr + 1
            curSum = curSum + rawEvents(i, 2)
            stepRes(pPtr, 1) = rawEvents(i, 1): stepRes(pPtr, 2) = curSum: pPtr = pPtr + 1
        Next i
        
        wsDest.Range("AV1:AW1").Value = Array("Plot Time (s)", "Plot Demand")
        wsDest.Range("AV2").Resize(UBound(stepRes, 1), 2).Value = stepRes
        wsDest.Range("AW2:AW" & UBound(stepRes, 1) + 1).NumberFormat = "0.0%"
    End If

    ' --- RECEPTION EFFICIENCY TABLE ---
    UpdateProgressBar 65, "Populating Reception Efficiency tables..."
    Dim effTable(1 To 5, 1 To 6) As Variant
    Dim txDict As Object: Set txDict = CreateObject("Scripting.Dictionary")
    
    wsDest.Range("Z2").Value = "Reception Efficiency Table"
    wsDest.Range("Z2").Font.Bold = True
    
    effTable(1, 1) = "TXs/SF"
    effTable(1, 2) = "SFs"
    effTable(1, 3) = "RXCNT"
    effTable(1, 4) = "MaxRX"
    effTable(1, 5) = "Rx Eff %"
    effTable(1, 6) = "C-V2X Eff %"
    
    For i = 1 To UBound(dataArr, 1)
        Dim txCount As Long: txCount = dataArr(i, colTXperSFN)
        If Not txDict.Exists(txCount) Then txDict.Add txCount, Array(0, 0, 0, 0)
        Dim dValTemp As Variant: dValTemp = txDict(txCount)
        dValTemp(0) = dValTemp(0) + 1
        dValTemp(1) = dValTemp(1) + CDbl(dataArr(i, colRxC))
        dValTemp(2) = dValTemp(2) + sfnGroupRx(dataArr(i, colSFN))
        dValTemp(3) = dValTemp(3) + (numRxMax - 1)
        txDict(txCount) = dValTemp
    Next i
    
    Dim rIdx As Integer: rIdx = 2: Dim txKey As Variant
    For Each txKey In txDict.Keys
        If rIdx <= 5 Then
            Dim dVal As Variant: dVal = txDict(txKey)
            effTable(rIdx, 1) = txKey
            effTable(rIdx, 2) = dVal(0)
            effTable(rIdx, 3) = dVal(1)
            effTable(rIdx, 4) = dVal(0) * (numRxMax - 1)
            If effTable(rIdx, 4) > 0 Then
                effTable(rIdx, 5) = (dVal(1) / effTable(rIdx, 4))
                Dim cvM As Double: cvM = dVal(0) * (numRxMax - txKey) * txKey
                If cvM > 0 Then effTable(rIdx, 6) = (dVal(2) / cvM)
            End If
            rIdx = rIdx + 1
        End If
    Next txKey
    
    ' Output values
    wsDest.Range("Z3:AE7").Value = effTable
    
    ' Apply Light Red Highlights, Centering, & Borders to Table Range (Z3:AE7)
    With wsDest.Range("Z3:AE3") ' Header Row styling
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(240, 128, 128) ' Light Coral / Darker Rose
        .HorizontalAlignment = xlCenter      ' Centered column headers
    End With
    
    With wsDest.Range("Z4:AE7") ' Data Rows styling
        .Interior.Color = RGB(255, 204, 204) ' Classic Soft Red Fill
    End With
    
    ' Custom Column Formatting & Alignment
    With wsDest.Range("Z4:Z7")
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter      ' Center values in Col 1 (TXs/SF)
    End With
    
    ' Format Columns 2, 3, and 4 as Integers and Right-Justify
    With wsDest.Range("AA4:AC7")
        .NumberFormat = "0"
        .HorizontalAlignment = xlRight       ' Right-justify SFs, RXCNT, and MaxRX
    End With
    
    ' Format Right Columns (Percentages)
    With wsDest.Range("AD4:AE7")
        .NumberFormat = "0.00%"
    End With
    
    ' Apply clean border grid lines
    With wsDest.Range("Z3:AE7").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 190, 190)
    End With
    
    ' --- PLOT 1: EFFICIENCY / LOAD / DEMAND ---
    UpdateProgressBar 80, "Generating Primary Time Series Plots..."
    Set chtObjTime = wsDest.ChartObjects.Add(350, 10, 850, 400)
    With chtObjTime.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        
        ' Series 1: Standard Rx Efficiency (Col AJ)
        Dim sc1 As Series: Set sc1 = .SeriesCollection.NewSeries
        sc1.Name = "Rx Eff": sc1.XValues = wsDest.Range("AI2:AI" & winCount): sc1.Values = wsDest.Range("AJ2:AJ" & winCount)
        
        ' Series 2: C-V2X Rx Efficiency (Col AK)
        Dim sc2 As Series: Set sc2 = .SeriesCollection.NewSeries
        sc2.Name = "C-V2X Rx Eff": sc2.XValues = wsDest.Range("AI2:AI" & winCount): sc2.Values = wsDest.Range("AK2:AK" & winCount)
        
        ' Series 3: Channel Load (Col AL)
        Dim sc3 As Series: Set sc3 = .SeriesCollection.NewSeries
        sc3.Name = "Load": sc3.XValues = wsDest.Range("AI2:AI" & winCount): sc3.Values = wsDest.Range("AL2:AL" & winCount)
        
        ' Optional Series 4: Staircase Application Demand
        If eCount > 0 Then
            Dim scDem As Series: Set scDem = .SeriesCollection.NewSeries
            scDem.Name = "Demand": scDem.XValues = wsDest.Range("AV2:AV" & UBound(stepRes, 1) + 1): scDem.Values = wsDest.Range("AW2:AW" & UBound(stepRes, 1) + 1)
            scDem.Format.Line.ForeColor.RGB = RGB(0, 112, 192): scDem.Format.Line.Weight = 1.5
        End If
        
        ' ===================================================================================
        ' deGPS PLOTTING LOGIC (Using Column 3 with Distinct Line Colors & Legended Names)
        ' ===================================================================================
        Dim hasGPSData As Boolean: hasGPSData = False
        
        If Not deGPSTable Is Nothing Then
            If Not (deGPSTable.DataBodyRange Is Nothing) Then
                deGPSData = deGPSTable.DataBodyRange.Value
                hasGPSData = True
            End If
        End If
        
        If hasGPSData Then
            Dim gpsRow As Long
            Dim plotMinX As Double: plotMinX = tStartVal / 1000
            Dim plotMaxX As Double: plotMaxX = tStopVal / 1000
            
            ' Color Palette for unique vertical lines (Red, Orange, Blue, Purple, Green, Pink)
            Dim lineColors() As Variant
            lineColors = Array(RGB(220, 0, 0), _
                               RGB(237, 125, 49), _
                               RGB(46, 117, 182), _
                               RGB(112, 48, 160), _
                               RGB(112, 173, 71), _
                               RGB(219, 48, 105))
            
            For gpsRow = 1 To UBound(deGPSData, 1)
                ' Column 3 holds the elapsed time value directly (seconds)
                Dim tGpsEvent As Double: tGpsEvent = CDbl(deGPSData(gpsRow, 3))
                
                ' Plot vertical line if it falls inside our active X-axis window limits
                If tGpsEvent >= plotMinX And tGpsEvent <= plotMaxX Then
                    Dim gpsSer As Series: Set gpsSer = .SeriesCollection.NewSeries
                    gpsSer.Name = "SYNC LOSS(" & gpsRow & ")"
                    gpsSer.XValues = Array(tGpsEvent, tGpsEvent)
                    gpsSer.Values = Array(0, 1)
                    
                    ' Assign unique color from palette cycling based on index
                    Dim colorIdx As Long: colorIdx = (gpsRow - 1) Mod (UBound(lineColors) + 1)
                    
                    With gpsSer.Format.Line
                        .ForeColor.RGB = lineColors(colorIdx)
                        .Weight = 1.75
                        .DashStyle = msoLineDash
                    End With
                    gpsSer.MarkerStyle = xlMarkerStyleNone
                End If
            Next gpsRow
        End If
        ' ===================================================================================
        
        .HasTitle = True
        .ChartTitle.Text = "Rx Efficiency / Load / Demand versus Time"
        
        ' Configure Axes and Labels
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Time (s)"
        End With
        With .Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 1
            .TickLabels.NumberFormat = "0%"
            .HasTitle = True
            .AxisTitle.Text = "Rx Efficiency / Load / Demand (%)"
        End With
        
        ' Enable horizontal legend below the h-axis
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' Apply Heatmap Colors to the Line Points
        ApplyHeatmap126 sc1, resultsArr, 2, True  ' Heatmap color for Rx Eff
        ApplyHeatmap126 sc2, resultsArr, 3, True  ' Heatmap color for C-V2X Rx Eff
        ApplyHeatmap126 sc3, resultsArr, 4, False ' Heatmap color for Channel Load
    End With

    ' --- PLOT 2: EFFICIENCY vs LOAD (No Legend) ---
    UpdateProgressBar 90, "Generating Load versus Efficiency scatter plot..."
    Set chtObjLoad = wsDest.ChartObjects.Add(350, 420, 850, 450)
    With chtObjLoad.Chart
        .ChartType = xlXYScatter
        Dim sc4 As Series: Set sc4 = .SeriesCollection.NewSeries
        sc4.XValues = wsDest.Range("AL2:AL" & winCount): sc4.Values = wsDest.Range("AJ2:AJ" & winCount)
        
        .HasTitle = True
        .ChartTitle.Text = "Rx Efficiency versus Load"
        .HasLegend = False
        
        ' Configure Axes and Labels
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "0%"
            .HasTitle = True
            .AxisTitle.Text = "Channel Load"
        End With
        With .Axes(xlValue)
            .MaximumScale = 1
            .TickLabels.NumberFormat = "0%"
            .HasTitle = True
            .AxisTitle.Text = "Rx Efficiency"
        End With
        
        ApplyHeatmap126 sc4, resultsArr, 4, False
    End With

    ' --- PLOT 3: FILTERED TX per SF (No Legend) ---
    UpdateProgressBar 95, "Generating TX allocation scatter plots..."
    Set chtObjTX = wsDest.ChartObjects.Add(wsDest.Range("Y11").Left, wsDest.Range("Y11").Top, 480, 375)
    With chtObjTX.Chart
        .ChartType = xlXYScatter
        Dim sTX As Series: Set sTX = .SeriesCollection.NewSeries
        sTX.XValues = wsDest.Range("AI2:AI" & winCount): sTX.Values = wsDest.Range("AM2:AM" & winCount)
        
        .HasTitle = True
        .ChartTitle.Text = "Filtered TX per Subframe"
        .HasLegend = False
        
        ' Configure Axes and Labels
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Time (s)"
        End With
        With .Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = finalYMax
            .HasTitle = True
            .AxisTitle.Text = "TX per Subframe"
        End With
        
        ApplyTXHeatmap sTX, wsDest.Range("AM2:AM" & winCount).Value
    End With
    
    wsDest.Activate
    
    ' Disable Cover / Restore Screen Settings
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    
    ' Clear the status bar
    Application.StatusBar = False
    
    ' If run within an automated pipeline, return elapsed duration back to the logTable dictionary
    Dim tElapsed As Double: tElapsed = MicroTimer() - tStart
    If Not logTable Is Nothing Then
        logTable("GenerateLoadRxEfficiencyAnalysis") = tElapsed
    End If
End Sub

' ===========================================================================================
' PRIVATE UTILITIES: PERFORMANCE SWITCH & PROGRESS BAR
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
' HEATMAP HELPERS (Maintained verbatim for graphic accuracy)
' ===========================================================================================
Sub ApplyTXHeatmap(ser As Series, valArr As Variant)
    Dim p As Long, v As Double, r As Integer, g As Integer
    Dim pCount As Long: pCount = IIf(ser.Points.count > 2000, 2000, ser.Points.count)
    
    For p = 1 To pCount
        v = valArr(p, 1)
        If v <= 1 Then
            r = 0: g = 255
        ElseIf v <= 1.5 Then
            r = 255: g = 255
        ElseIf v <= 2 Then
            r = 255: g = 165
        Else
            r = 255: g = 0
        End If
        
        With ser.Points(p)
            .MarkerStyle = xlMarkerStyleCircle: .MarkerSize = 4
            .MarkerBackgroundColor = RGB(r, g, 0): .MarkerForegroundColor = RGB(r, g, 0)
        End With
        If p Mod 100 = 0 Then DoEvents
    Next p
End Sub

Sub ApplyHeatmap126(ser As Series, data As Variant, colIdx As Integer, invert As Boolean)
    Dim p As Long, ptVal As Double, scaledVal As Double, r As Integer, g As Integer
    Dim pCount As Long: pCount = ser.Points.count
    
    For p = 1 To pCount
        ptVal = data(p, colIdx)
        
        If invert Then
            scaledVal = (ptVal - 0.5) / 0.5
        Else
            scaledVal = 1 - ptVal
        End If
        
        scaledVal = WorksheetFunction.Median(0, 1, scaledVal)
        
        If scaledVal <= 0.5 Then
            r = 255: g = Int(255 * (scaledVal * 2))
        Else
            r = Int(255 * (2 - (scaledVal * 2))): g = 255
        End If
        
        With ser.Points(p)
            .MarkerStyle = xlMarkerStyleCircle: .MarkerSize = 5
            .MarkerBackgroundColor = RGB(r, g, 0): .MarkerForegroundColor = RGB(r, g, 0)
        End With
        If p Mod 100 = 0 Then DoEvents
    Next p
End Sub



