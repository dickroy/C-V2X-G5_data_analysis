Attribute VB_Name = "mod_08_TxGapPerApp"
' ===========================================================================================
' Module: Tx_Gap_Per_App_Analysis
' Version: 3.0.0 (STABLE - ELIMINATED .NET DEPENDENCY & ADDED PERFORMANCE OPTIMIZATIONS)
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
' WRAPPER FOR MANUAL EXECUTION
' ===========================================================================================
Sub Run_GenerateTxGapPerAppAnalysis()
    Dim startTime As Double: startTime = MicroTimer()
    
    ' Call core macro with pipeline parameters set to manual state
    Call GenerateTxGapPerAppAnalysis(Nothing)
    
    ' Output execution timing only on manual runs
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "Tx Gap per App Analysis Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT
' ===========================================================================================
Sub GenerateTxGapPerAppAnalysis(Optional ByRef logTable As Object = Nothing)
    Dim tStartExec As Double: tStartExec = MicroTimer() ' Start Performance Tracking

    Dim wsSrc As Worksheet, wsDest As Worksheet, wsApp As Worksheet
    Dim targetTable As ListObject, appTable As ListObject
    Dim dataArr As Variant, appArr As Variant
    Dim i As Long, j As Long
    Dim msgID As String, lastMsgID As String
    Dim txTime As Double, lastTxTime As Double, delta As Double
    Dim maxDelta As Double: maxDelta = 120
    
    ' Trackers
    Dim txGaps As Object: Set txGaps = CreateObject("Scripting.Dictionary")
    Dim binTracker As Object: Set binTracker = CreateObject("Scripting.Dictionary")
    Dim binTrackerLarge As Object: Set binTrackerLarge = CreateObject("Scripting.Dictionary")
    
    Dim countUnderEqual120 As Long: countUnderEqual120 = 0
    Dim countOver120 As Long: countOver120 = 0
    Dim totalTTIs As Long
    Dim chtObj As ChartObject
    Dim lc As ListColumn, cleanName As String
    
    ' Header location variables
    Dim colMID As Long: colMID = 0
    Dim colTxQ As Long: colTxQ = 0
    Dim colMID_App As Long: colMID_App = 0
    Dim colSID As Long: colSID = 0
    Dim colADU As Long: colADU = 0
    Dim colTTI As Long: colTTI = 0
    
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

    ' --- VALIDATION BLOCK ---
    Dim hasValidEntry As Boolean: hasValidEntry = False
    
    ' Map validation columns in AppParams
    For Each lc In appTable.ListColumns
        cleanName = UCase(Trim(lc.Name))
        Select Case cleanName
            Case "STATION ID", "STATION_ID": colSID = lc.Index
            Case "ADU SIZE (B)", "ADU_SIZE": colADU = lc.Index
            Case "TTI(MS)", "TTI": colTTI = lc.Index
            Case "APP_ID", "MSG_ID", "IVI_ID": colMID_App = lc.Index
        End Select
    Next lc

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

    ' Dynamically find Column Indices in Results Table (Case-insensitive)
    For j = 1 To targetTable.ListColumns.count
        cleanName = UCase(Trim(targetTable.ListColumns(j).Name))
        If cleanName = "APP_ID" Or cleanName = "MSG_ID" Or cleanName = "IVI_ID" Then
            colMID = j
        End If
        If cleanName = "TXQTIME" Then
            colTxQ = j
        End If
    Next j

    If colMID = 0 Or colTxQ = 0 Then
        CoverScreen False
        Application.StatusBar = False
        MsgBox "Critical Error: Could not find App ID or TXQTIME columns in results.", vbCritical
        Exit Sub
    End If

    ' 2. FORCE SORT BEFORE PROCESSING
    With targetTable.Sort
        .SortFields.Clear
        .SortFields.Add key:=targetTable.ListColumns(colMID).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add key:=targetTable.ListColumns(colTxQ).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Prepare Destination Sheet
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Tx Gap per App Analysis")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "Tx Gap per App Analysis"
    End If
    wsDest.Cells.Clear
    For Each chtObj In wsDest.ChartObjects
        chtObj.Delete
    Next chtObj
    On Error GoTo 0
    
    dataArr = targetTable.Range.Value

    ' 3. Initialize Bins
    For i = 0 To 300
        binTracker(Round(90# + (i * 0.1), 1)) = 0
    Next i

    ' 4. Process Data
    UpdateProgressBar 45, "Calculating Tx Gaps and registering histogram bins..."
    lastMsgID = ""
    For i = 2 To UBound(dataArr, 1)
        msgID = CStr(dataArr(i, colMID))
        If IsNumeric(dataArr(i, colTxQ)) And dataArr(i, colTxQ) <> "" Then
            txTime = CDbl(dataArr(i, colTxQ))
            
            If msgID = lastMsgID Then
                delta = txTime - lastTxTime
                If delta > 0 Then
                    If delta > maxDelta Then
                        maxDelta = delta
                    End If
                    
                    If delta <= 120 Then
                        countUnderEqual120 = countUnderEqual120 + 1
                    Else
                        countOver120 = countOver120 + 1
                    End If
                    
                    ' Safe fallback replacement for .NET ArrayList using dynamic arrays in collection
                    If Not txGaps.Exists(msgID) Then
                        Dim tempArr() As Double
                        ReDim tempArr(0 To 0)
                        tempArr(0) = delta
                        txGaps.Add msgID, tempArr
                    Else
                        Dim extArr() As Double
                        extArr = txGaps(msgID)
                        ReDim Preserve extArr(0 To UBound(extArr) + 1)
                        extArr(UBound(extArr)) = delta
                        txGaps(msgID) = extArr
                    End If
                    
                    Dim targetKey As Double
                    If delta < 90 Then
                        targetKey = 90#
                    ElseIf delta >= 120 Then
                        targetKey = 120#
                    Else
                        targetKey = Round(Int(delta * 10) / 10, 1)
                    End If
                    binTracker(targetKey) = binTracker(targetKey) + 1
                    
                    If delta > 120 Then
                        Dim lKey As Long: lKey = Int(delta)
                        binTrackerLarge(lKey) = binTrackerLarge(lKey) + 1
                    End If
                End If
            Else
                lastMsgID = msgID
            End If
            lastTxTime = txTime
        End If
    Next i
    
    totalTTIs = countUnderEqual120 + countOver120

    ' 5. Output Raw Data
    UpdateProgressBar 65, "Outputting raw delta structures..."
    Dim dataStartCol As Long: dataStartCol = 18
    Dim curCol As Long: curCol = dataStartCol
    Dim key As Variant
    For Each key In txGaps.Keys
        wsDest.Cells(1, curCol).Value = "MSG_" & key
        Dim items() As Double: items = txGaps(key)
        Dim boundSize As Long: boundSize = UBound(items) + 1
        
        If boundSize > 0 Then
            Dim outArr() As Variant: ReDim outArr(1 To boundSize, 1 To 1)
            For j = 0 To UBound(items)
                outArr(j + 1, 1) = items(j)
            Next j
            wsDest.Cells(2, curCol).Resize(boundSize, 1).Value = outArr
        End If
        curCol = curCol + 1
    Next key

    ' 6. Frequency Tables (Optimized via Array Bulk Write)
    Dim binTableCol As Long: binTableCol = curCol + 2
    wsDest.Cells(1, binTableCol).Value = "Bin (ms)"
    wsDest.Cells(1, binTableCol + 1).Value = "Freq (90-120)"
    
    Dim bulkFreq(1 To 302, 1 To 2) As Variant
    For i = 0 To 300
        Dim cb As Double: cb = Round(90# + (i * 0.1), 1)
        bulkFreq(i + 1, 1) = cb
        bulkFreq(i + 1, 2) = binTracker(cb)
    Next i
    wsDest.Cells(2, binTableCol).Resize(301, 2).Value = bulkFreq
    
    Dim binColLarge As Long: binColLarge = binTableCol + 3
    wsDest.Cells(1, binColLarge).Value = "Bin (>120ms)"
    wsDest.Cells(1, binColLarge + 1).Value = "Freq (>120)"
    
    Dim maxLargeIndex As Long: maxLargeIndex = Int(maxDelta)
    Dim largeCount As Long: largeCount = maxLargeIndex - 120 + 1
    
    If largeCount > 0 Then
        Dim bulkLarge() As Variant: ReDim bulkLarge(1 To largeCount, 1 To 2)
        Dim rOff As Long: rOff = 1
        For i = 120 To maxLargeIndex
            bulkLarge(rOff, 1) = i
            If binTrackerLarge.Exists(i) Then
                bulkLarge(rOff, 2) = binTrackerLarge(i)
            Else
                bulkLarge(rOff, 2) = 0
            End If
            rOff = rOff + 1
        Next i
        wsDest.Cells(2, binColLarge).Resize(largeCount, 2).Value = bulkLarge
    End If

    ' 7. Charts
    UpdateProgressBar 85, "Drawing visualization components..."
    Set chtObj = wsDest.ChartObjects.Add(10, 10, 750, 350)
    With chtObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsDest.Range(wsDest.Cells(2, binTableCol + 1), wsDest.Cells(302, binTableCol + 1))
        .SeriesCollection(1).XValues = wsDest.Range(wsDest.Cells(2, binTableCol), wsDest.Cells(302, binTableCol))
        .HasTitle = True
        .ChartTitle.Text = "Transmission Time Interval (ms) (All Apps)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Actual TTI (ms)"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Frequency"
        .ChartGroups(1).GapWidth = 50
        .Axes(xlCategory).TickLabelSpacing = 10
        .Axes(xlCategory).TickLabels.Orientation = 45
        .HasLegend = False
    End With

    If largeCount > 0 Then
        Set chtObj = wsDest.ChartObjects.Add(10, 380, 750, 350)
        With chtObj.Chart
            .ChartType = xlColumnClustered
            .SetSourceData Source:=wsDest.Range(wsDest.Cells(2, binColLarge + 1), wsDest.Cells(largeCount + 1, binColLarge + 1))
            .SeriesCollection(1).XValues = wsDest.Range(wsDest.Cells(2, binColLarge), wsDest.Cells(largeCount + 1, binColLarge))
            .HasTitle = True
            .ChartTitle.Text = "Transmission Time Interval (> 120ms) (All Apps)"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Actual TTI (ms)"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Frequency"
            .ChartGroups(1).GapWidth = 50
            .HasLegend = False
        End With
    End If

    ' 8. Summary Table
    Dim sumRow As Long: sumRow = 60
    With wsDest
        .Cells(sumRow, 1).Value = "Actual TTI Range"
        .Cells(sumRow, 2).Value = "Number of TTIs"
        .Cells(sumRow, 3).Value = "% of Total"
        .Cells(sumRow + 1, 1).Value = "<=120 (ms)"
        .Cells(sumRow + 1, 2).Value = countUnderEqual120
        .Cells(sumRow + 2, 1).Value = ">120 (ms)"
        .Cells(sumRow + 2, 2).Value = countOver120
        If totalTTIs > 0 Then
            .Cells(sumRow + 1, 3).Value = countUnderEqual120 / totalTTIs
            .Cells(sumRow + 2, 3).Value = countOver120 / totalTTIs
        End If
        .Range(.Cells(sumRow, 1), .Cells(sumRow, 3)).Font.Bold = True
        .Range(.Cells(sumRow + 1, 3), .Cells(sumRow + 2, 3)).NumberFormat = "0.00%"
        .Range(.Cells(sumRow, 1), .Cells(sumRow + 2, 3)).Borders.LineStyle = xlContinuous
    End With

    ' 9. Highlight Outliers
    If curCol > dataStartCol Then
        Dim rawDataRange As Range
        Set rawDataRange = wsDest.Range(wsDest.Cells(2, dataStartCol), wsDest.Cells(UBound(dataArr, 1), curCol - 1))
        rawDataRange.FormatConditions.Delete
        rawDataRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=120"
        rawDataRange.FormatConditions(1).Font.Color = RGB(255, 0, 0)
        rawDataRange.FormatConditions(1).Interior.Color = RGB(255, 230, 230)
    End If

    wsDest.Columns.AutoFit
    wsDest.Activate

    ' Clean exit routine
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    Application.StatusBar = False
    
    Dim tElapsedTotal As Double: tElapsedTotal = MicroTimer() - tStartExec
    If Not logTable Is Nothing Then
        logTable("GenerateTxGapPerAppAnalysis") = tElapsedTotal
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

