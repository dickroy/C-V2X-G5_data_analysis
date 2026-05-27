Attribute VB_Name = "mod_07_RxGapPerApp"
' ===========================================================================================
' Module: Rx_Gap_Per_App_Analysis
' Version: 2.4.2-INTEGRATED (STABLE COMPILATION - ALL SYNTAX TYPOS RESOLVED)
' ===========================================================================================

Option Explicit

' Native high-precision timing API (supports 64-bit and 32-bit Excel environments)
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
Sub Run_GenerateRxGapPerAppAnalysis()
    Dim startTime As Double: startTime = MicroTimer()
    
    ' Call core macro with pipeline parameters set to manual state
    Call GenerateRxGapPerAppAnalysis(Nothing)
    
    ' Output execution timing only on manual runs
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "Rx Gap per App Analysis Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT (Safe for automated pipelines calling GenerateRxGapPerAppAnalysis(Nothing))
' ===========================================================================================
Sub GenerateRxGapPerAppAnalysis(Optional ByRef logTable As Object = Nothing)
    Dim tStartExec As Double: tStartExec = MicroTimer() ' Start Performance Tracking

    Dim wsSrc As Worksheet, wsDest As Worksheet, wsApp As Worksheet
    Dim targetTable As ListObject, appTable As ListObject
    Dim dataArr As Variant, appArr As Variant
    Dim i As Long, j As Long, col As Long
    Dim msgID As String, lastMsgID As String
    Dim currentGap As Long
    Dim rxColIdx As Variant, colMID As Long, colMID_App As Long, colAC_App As Long
    Dim rxCols As New Collection
    Dim gapTracker As Object: Set gapTracker = CreateObject("Scripting.Dictionary")
    Dim acTracker As Object: Set acTracker = CreateObject("Scripting.Dictionary")
    Dim acMap As Object: Set acMap = CreateObject("Scripting.Dictionary")
    Dim chtObj As ChartObject
    Dim lc As ListColumn, cleanName As String
    Dim seriesName As String
    
    ' Toggle Excel engine optimizations on & push starting progress
    CoverScreen True
    UpdateProgressBar 5, "Locating data sources and configuration parameters..."
    
    ' 1. INITIALIZE TABLES
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    Set wsApp = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    Set appTable = wsApp.ListObjects("AppParams")
    On Error GoTo 0

    If targetTable Is Nothing Or appTable Is Nothing Then
        CoverScreen False
        Application.StatusBar = False
        MsgBox "Critical Error: Required tables not found.", vbCritical
        Exit Sub
    End If

    ' --- NEW VALIDATION BLOCK ---
    Dim hasValidEntry As Boolean: hasValidEntry = False
    Dim colSID As Long, colADU As Long, colTTI As Long
    
    ' Map validation columns
    For Each lc In appTable.ListColumns
        cleanName = UCase(Trim(lc.Name))
        Select Case cleanName
            Case "STATION ID", "STATION_ID": colSID = lc.Index
            Case "ADU SIZE (B)", "ADU_SIZE": colADU = lc.Index
            Case "TTI(MS)", "TTI": colTTI = lc.Index
            Case "APP_ID", "MSG_ID", "IVI_ID": colMID_App = lc.Index
        End Select
    Next lc

    ' Verify columns exist and check for at least one complete row
    If appTable.DataBodyRange Is Nothing Then
        CoverScreen False
        Application.StatusBar = False
        MsgBox "AppParams table is empty.", vbExclamation
        Exit Sub
    End If

    appArr = appTable.DataBodyRange.Value
    If colMID_App > 0 And colSID > 0 And colADU > 0 And colTTI > 0 Then
        For i = 1 To UBound(appArr, 1)
            If Trim(CStr(appArr(i, colMID_App))) <> "" And _
               Trim(CStr(appArr(i, colSID))) <> "" And _
               Trim(CStr(appArr(i, colADU))) <> "" And _
               Trim(CStr(appArr(i, colTTI))) <> "" Then
                hasValidEntry = True
                Exit For
            End If
        Next i
    End If

    If Not hasValidEntry Then
        CoverScreen False
        Application.StatusBar = False
        MsgBox "Error: AppParams table must contain at least one valid entry " & vbCrLf & _
               "(App_ID, Station ID, ADU Size, and TTI are required).", vbCritical
        Exit Sub
    End If
    ' --- END VALIDATION BLOCK ---

    UpdateProgressBar 20, "Dynamic header mapping and applying target sorts..."

    ' 2. DYNAMIC HEADER MAPPING
    colMID = 0: colAC_App = 0
    
    For Each lc In appTable.ListColumns
        cleanName = UCase(Trim(lc.Name))
        If cleanName = "AC" Or cleanName = "ACCESS_CLASS" Then
            colAC_App = lc.Index
        End If
    Next lc

    For j = 1 To targetTable.ListColumns.count
        cleanName = UCase(Trim(targetTable.ListColumns(j).Name))
        If cleanName = "APP_ID" Or cleanName = "MSG_ID" Or cleanName = "IVI_ID" Then
            colMID = j
        End If
    Next j

    If colMID = 0 Or colMID_App = 0 Then
        CoverScreen False
        Application.StatusBar = False
        MsgBox "Critical Error: Could not find an ID column in tables.", vbCritical
        Exit Sub
    End If

    ' 3. FORCE SORT: GROUP BY APP FOR GAP CALC
    With targetTable.Sort
        .SortFields.Clear
        .SortFields.Add key:=targetTable.ListColumns(colMID).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add key:=targetTable.ListColumns("TXQTIME").Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    ' 4. PREPARE DESTINATION SHEET
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Rx Gap per App Analysis")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "Rx Gap per App Analysis"
    End If
    wsDest.Cells.Clear
    For Each chtObj In wsDest.ChartObjects
        chtObj.Delete
    Next chtObj
    On Error GoTo 0
    
    dataArr = targetTable.Range.Value
    
    For i = 1 To UBound(appArr, 1)
        If colAC_App > 0 Then
            acMap(CStr(appArr(i, colMID_App))) = "AC " & appArr(i, colAC_App)
        Else
            acMap(CStr(appArr(i, colMID_App))) = "AC Unknown"
        End If
    Next i

    For j = 1 To UBound(dataArr, 2)
        cleanName = UCase(Trim(dataArr(1, j)))
        If InStr(1, cleanName, "RXTIME") > 0 And InStr(1, cleanName, "TX_ID") = 0 Then
            rxCols.Add j
        End If
    Next j

    ' 5. PROCESS GAPS (WITH OVERFLOW ACCUMULATOR AT >10)
    UpdateProgressBar 45, "Scanning Rx events and tracking contiguous packet drops..."
    
    ' Bins are strictly mapped 0 to 11 (11 represents the ">10" overflow category)
    Dim binKey As Long
    For Each rxColIdx In rxCols
        currentGap = 0: lastMsgID = ""
        For i = 2 To UBound(dataArr, 1)
            msgID = CStr(dataArr(i, colMID))
            
            If msgID <> lastMsgID Then
                currentGap = 0
                lastMsgID = msgID
                If Not gapTracker.Exists(msgID) Then
                    gapTracker.Add msgID, CreateObject("Scripting.Dictionary")
                End If
            End If
            
            If dataArr(i, rxColIdx) = "" Then
                currentGap = currentGap + 1
            Else
                If currentGap > 0 Then
                    ' Clamp anything above 10 into the overflow slot (index 11)
                    binKey = IIf(currentGap > 10, 11, currentGap)
                    gapTracker(msgID)(binKey) = gapTracker(msgID)(binKey) + 1
                Else
                    gapTracker(msgID)(0) = gapTracker(msgID)(0) + 1
                End If
                currentGap = 0
            End If
        Next i
    Next rxColIdx

    If gapTracker.count = 0 Then
        GoTo CleanRestore
    End If

    ' 6. OUTPUT DATA TABLES (CAPPED AT 10 + OVERFLOW COLUMN)
    UpdateProgressBar 65, "Populating gap structures and Access Category (AC) matrices..."
    Dim tableStartCol As Integer: tableStartCol = 14
    wsDest.Cells(1, tableStartCol).Value = "APP_ID"
    
    ' Define explicit column headers
    For j = 0 To 10
        wsDest.Cells(1, tableStartCol + j + 1).Value = j
    Next j
    wsDest.Cells(1, tableStartCol + 12).Value = ">10" ' Explicit column 12 for the terminal overflow bin
    
    Dim rowIdx As Long: rowIdx = 2
    Dim totals(0 To 11) As Double ' Indices 0 to 10 represent 0 to 10; Index 11 represents ">10"
    Dim key As Variant, curAC As String
    
    For Each key In gapTracker.Keys
        wsDest.Cells(rowIdx, tableStartCol).Value = key
        If acMap.Exists(CStr(key)) Then
            curAC = acMap(CStr(key))
        Else
            curAC = "Unknown"
        End If
        
        If Not acTracker.Exists(curAC) Then
            acTracker.Add curAC, CreateObject("Scripting.Dictionary")
        End If
        
        For j = 0 To 11
            If gapTracker(key).Exists(j) Then
                Dim val As Long: val = gapTracker(key)(j)
                wsDest.Cells(rowIdx, tableStartCol + j + 1).Value = val
                totals(j) = totals(j) + val
                acTracker(curAC)(j) = acTracker(curAC)(j) + val
            End If
        Next j
        rowIdx = rowIdx + 1
    Next key
    
    wsDest.Cells(rowIdx, tableStartCol).Value = "TOTALS"
    wsDest.Cells(rowIdx, tableStartCol + 1).Resize(1, 12).Value = totals
    wsDest.Range(wsDest.Cells(rowIdx, tableStartCol), wsDest.Cells(rowIdx, tableStartCol + 12)).Font.Bold = True

    ' AC Summary Table
    Dim startColAC As Integer: startColAC = tableStartCol + 14
    wsDest.Cells(1, startColAC).Value = "Access Category (AC)"
    For j = 0 To 10
        wsDest.Cells(1, startColAC + j + 1).Value = j
    Next j
    wsDest.Cells(1, startColAC + 12).Value = ">10"
    
    Dim acRow As Long: acRow = 2
    For Each key In acTracker.Keys
        wsDest.Cells(acRow, startColAC).Value = key
        For j = 0 To 11
            If acTracker(key).Exists(j) Then
                wsDest.Cells(acRow, startColAC + j + 1).Value = acTracker(key)(j)
            End If
        Next j
        acRow = acRow + 1
    Next key

    ' 7. CHARTING (1 to 10 + ">10", Category 0 Omitted)
    UpdateProgressBar 85, "Generating analytical distribution charts..."
    
    ' Safeguard against cases where maxGap is 0 (cannot plot 1 to 0)
    If gapTracker.count > 0 Then
        ' Chart 1: Totals (Plots from index 1 through 10 and terminal ">10" bin)
        Set chtObj = wsDest.ChartObjects.Add(wsDest.Cells(1, 1).Left, wsDest.Cells(1, 1).Top, 455, 280)
        With chtObj.Chart
            .ChartType = xlColumnClustered
            With .SeriesCollection.NewSeries
                ' Category headers start at tableStartCol + 2 (which holds '1') up to tableStartCol + 12 (which holds '>10')
                .XValues = wsDest.Range(wsDest.Cells(1, tableStartCol + 2), wsDest.Cells(1, tableStartCol + 12))
                .Values = wsDest.Range(wsDest.Cells(rowIdx, tableStartCol + 2), wsDest.Cells(rowIdx, tableStartCol + 12))
                .Name = "Total Gaps"
                .HasDataLabels = True
            End With
            .HasTitle = True
            .ChartTitle.Text = "Consecutive Missed Receptions (All Apps)"
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Number of Consecutive Missed Receptions"
            End With
            .HasLegend = False
        End With

        ' Chart 2: AC Breakdown (Plots from index 1 through 10 and terminal ">10" bin with color formatting)
        ' DATALABELS REMOVED TO PREVENT OVERLAPPING READABILITY ISSUES
        Set chtObj = wsDest.ChartObjects.Add(wsDest.Cells(1, 1).Left, wsDest.Cells(21, 1).Top, 455, 280)
        With chtObj.Chart
            .ChartType = xlColumnClustered
            For i = 2 To acRow - 1
                seriesName = wsDest.Cells(i, startColAC).Value
                With .SeriesCollection.NewSeries
                    .Name = seriesName
                    ' Category headers start at startColAC + 2 (which holds '1') up to startColAC + 12 (which holds '>10')
                    .XValues = wsDest.Range(wsDest.Cells(1, startColAC + 2), wsDest.Cells(1, startColAC + 12))
                    .Values = wsDest.Range(wsDest.Cells(i, startColAC + 2), wsDest.Cells(i, startColAC + 12))
                    
                    ' Data labels explicitly disabled for this crowded chart
                    .HasDataLabels = False
                    
                    ' Apply custom colors based on AC category (AC0 = RED, AC1 = YELLOW, AC2 = GREEN, AC3 = BLUE)
                    If InStr(1, seriesName, "0") > 0 Then
                        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)     ' RED
                    ElseIf InStr(1, seriesName, "1") > 0 Then
                        .Format.Fill.ForeColor.RGB = RGB(255, 255, 0)   ' YELLOW
                    ElseIf InStr(1, seriesName, "2") > 0 Then
                        .Format.Fill.ForeColor.RGB = RGB(0, 255, 0)     ' GREEN
                    ElseIf InStr(1, seriesName, "3") > 0 Then
                        .Format.Fill.ForeColor.RGB = RGB(0, 0, 255)     ' BLUE
                    End If
                End With
            Next i
            .HasTitle = True
            .ChartTitle.Text = "Consecutive Missed Receptions by AC"
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Number of Consecutive Missed Receptions"
            End With
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
        End With
    End If

CleanRestore:
    UpdateProgressBar 95, "Restoring default experimental sort criteria..."
    With targetTable.Sort
        .SortFields.Clear
        .SortFields.Add key:=targetTable.ListColumns("TXQTIME").Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Apply
    End With
    wsDest.Columns.AutoFit
    wsDest.Activate

    ' Clean exit routine
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    Application.StatusBar = False
    
    ' Return run duration if executed from a master macro pipeline
    Dim tElapsedTotal As Double: tElapsedTotal = MicroTimer() - tStartExec
    If Not logTable Is Nothing Then
        logTable("GenerateRxGapPerAppAnalysis") = tElapsedTotal
    End If
End Sub

' ===========================================================================================
' PRIVATE COVERSCREEN & ENGINE CONTROLLERS
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

