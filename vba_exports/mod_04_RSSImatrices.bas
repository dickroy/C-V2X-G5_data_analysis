Attribute VB_Name = "mod_04_RSSImatrices"
' ===========================================================================================
' Module: RSSI_Matrix_Analysis
' Version: 2.0.0 - O(N) SEGMENTED PROCESSING WITH PERFORMANCE & UI WRAPPERS
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
Sub Run_GenerateRSSIMatrices()
    Dim startTime As Double: startTime = MicroTimer()
    
    ' Call the core routine (passing Nothing for the logTable)
    Call GenerateRSSIMatrices(Nothing)
    
    ' Output execution timing only on manual runs
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "RSSI Matrix Generation Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT (With Progress bar, performance covers, and optimized extraction)
' ===========================================================================================
Sub GenerateRSSIMatrices(Optional ByRef logTable As Object = Nothing)
    Dim tStart As Double: tStart = MicroTimer()
    
    Dim wsSrc As Worksheet, wsDest As Worksheet, wsApp As Worksheet
    Dim targetTable As ListObject
    Dim dataArr As Variant
    Dim numRx As Integer, i As Long, j As Long, k As Long
    Dim txIdCol As Long
    Dim noteStartCol As Integer, noteEndCol As Integer
    
    ' 1. SETUP
    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    Set wsApp = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
    
    numRx = [Num_Rx_Stations]
    
    ' Enable Performance Mode
    CoverScreen True
    UpdateProgressBar 5, "Initializing RSSI Workspace..."
    
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("RSSI Analysis")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = "RSSI Analysis"
    End If
    On Error GoTo 0
    wsDest.Cells.Clear
    
    ' ---------------------------------------------------------------------------------------
    ' OPTIMIZATION STEP 1: SORT AND CAPTURE CONTIGUOUS DATA STRIPES
    ' ---------------------------------------------------------------------------------------
    UpdateProgressBar 10, "Optimizing table layout for fast processing..."
    
    ' Sort physically to get contiguous TX_IDs
    With targetTable.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=targetTable.ListColumns("TX_ID").DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    dataArr = targetTable.DataBodyRange.Value
    txIdCol = targetTable.ListColumns("TX_ID").Index
    
    ' Restore original sort immediately to keep pipeline clean
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
    End If
    
    ' ---------------------------------------------------------------------------------------
    ' OPTIMIZATION STEP 2: BUILD AN O(1) TRANSMITTER INDEX MAP IN MEMORY
    ' ---------------------------------------------------------------------------------------
    UpdateProgressBar 20, "Building memory-indexed maps..."
    Dim txStartRows As Object: Set txStartRows = CreateObject("Scripting.Dictionary")
    Dim txEndRows As Object: Set txEndRows = CreateObject("Scripting.Dictionary")
    
    Dim numRows As Long: numRows = UBound(dataArr, 1)
    Dim currentTX As String, lastTX As String: lastTX = ""
    
    For k = 1 To numRows
        currentTX = CStr(dataArr(k, txIdCol))
        If currentTX <> lastTX Then
            If lastTX <> "" Then txEndRows(lastTX) = k - 1
            txStartRows(currentTX) = k
            lastTX = currentTX
        End If
    Next k
    If lastTX <> "" Then txEndRows(lastTX) = numRows
    
    ' 2. INITIALIZE MATRICES
    Dim matMean() As Variant, matStd() As Variant, matCount() As Variant, matDiff() As Variant
    ReDim matMean(1 To numRx, 1 To numRx)
    ReDim matStd(1 To numRx, 1 To numRx)
    ReDim matCount(1 To numRx, 1 To numRx)
    ReDim matDiff(1 To numRx, 1 To numRx)
    
    ' 3. PROCESSING LOOP (Utilizing Segmented Boundaries & Fast Arrays)
    Dim lastPercent As Integer: lastPercent = 0
    Dim currentPercent As Integer
    Dim totalIter As Long: totalIter = numRx * numRx
    Dim currentIter As Long: currentIter = 0
    
    For j = 1 To numRx ' Receiver (Column)
        Dim rssiColName As String: rssiColName = "RSSI" & j
        Dim colIdx As Long: colIdx = targetTable.ListColumns(rssiColName).Index
        
        For i = 1 To numRx ' Transmitter (Row)
            currentIter = currentIter + 1
            
            ' Throttled UI Progress Updates
            currentPercent = 25 + Round((currentIter / totalIter) * 60)
            If currentPercent <> lastPercent Then
                UpdateProgressBar currentPercent, "Analyzing RSSI path: TX " & i & " -> RX " & j & "..."
                lastPercent = currentPercent
            End If
            
            Dim mW_Sum As Double: mW_Sum = 0
            Dim count As Long: count = 0
            
            Dim sRow As Long: sRow = 0
            Dim eRow As Long: eRow = 0
            
            ' Access boundary maps directly in constant O(1) CPU time
            If txStartRows.Exists(CStr(i)) Then
                sRow = txStartRows(CStr(i))
                eRow = txEndRows(CStr(i))
            End If
            
            If sRow > 0 Then
                Dim segmentLen As Long: segmentLen = eRow - sRow + 1
                Dim dbm_Arr() As Double: ReDim dbm_Arr(1 To segmentLen)
                
                ' Process only the rows matching Transmitter i (Perfect Linear Speedup)
                For k = sRow To eRow
                    Dim val As Variant: val = dataArr(k, colIdx)
                    If IsNumeric(val) And val < 0 Then
                        count = count + 1
                        mW_Sum = mW_Sum + (10 ^ (val / 10))
                        dbm_Arr(count) = CDbl(val)
                    End If
                Next k
                
                If count > 0 And i <> j Then
                    ' MEAN: Linear average (mW) -> back to dBm
                    matMean(i, j) = 10 * (Log(mW_Sum / count) / Log(10))
                    
                    ' STD DEV: Calculated on flat array
                    If count > 1 Then
                        ' Squeeze array size to match matching values exactly
                        Dim calcArr() As Double: ReDim calcArr(1 To count)
                        Dim n As Long
                        For n = 1 To count: calcArr(n) = dbm_Arr(n): Next n
                        matStd(i, j) = WorksheetFunction.StDev_S(calcArr)
                    Else
                        matStd(i, j) = 0
                    End If
                    matCount(i, j) = CLng(count)
                Else
                    matMean(i, j) = ""
                    matStd(i, j) = ""
                    matCount(i, j) = ""
                End If
            Else
                matMean(i, j) = ""
                matStd(i, j) = ""
                matCount(i, j) = ""
            End If
        Next i
    Next j
    
    ' 4. CALCULATE MEAN - TRANSPOSE
    For i = 1 To numRx: For j = 1 To numRx
        If IsNumeric(matMean(i, j)) And IsNumeric(matMean(j, i)) Then
            matDiff(i, j) = matMean(i, j) - matMean(j, i)
        Else
            matDiff(i, j) = ""
        End If
    Next j: Next i
    
    ' 5. OUTPUT DATA
    UpdateProgressBar 90, "Writing analysis matrices to sheet..."
    PrintMatrix wsDest, "Mean RSSI (dBm)", matMean, 1, numRx, "0.00"
    PrintMatrix wsDest, "StdDev (dB)", matStd, (numRx + 3), numRx, "0.00"
    PrintMatrix wsDest, "Mean - Transpose", matDiff, (numRx * 2 + 5), numRx, "0.00"
    PrintMatrix wsDest, "Packet Count", matCount, (numRx * 3 + 7), numRx, "0"

    ' 6. DYNAMIC METHODOLOGY NOTE PLACEMENT
    noteStartCol = numRx + 3
    noteEndCol = noteStartCol + 4
    
    With wsDest.Range(wsDest.Cells(2, noteStartCol), wsDest.Cells(6, noteEndCol))
        .Merge
        .Value = "METHODOLOGY NOTE:" & vbCrLf & _
                 "• Mean(RSSI) is calculated in linear space (mW) and converted back to dBm." & vbCrLf & _
                 "• StdDev(RSSI) is calculated directly on the raw dBm values."
        .Font.Italic = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .WrapText = True
        .Borders.Weight = xlMedium
        .Interior.Color = RGB(245, 245, 245)
    End With

    wsDest.Columns.AutoFit
    wsDest.Activate
    
    ' 7. TEARDOWN PERFORMANCE MODE
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    Application.StatusBar = False
    
    Dim tElapsed As Double: tElapsed = MicroTimer() - tStart
    If Not logTable Is Nothing Then
        logTable("GenerateRSSIMatrices") = tElapsed
    End If
End Sub

' ===========================================================================================
' PRINT MATRIX UTILITY
' ===========================================================================================
Sub PrintMatrix(ws As Worksheet, title As String, mat As Variant, startRow As Long, Size As Integer, numFormat As String)
    ws.Cells(startRow, 1).Value = title
    ws.Cells(startRow, 1).Font.Bold = True
    Dim i As Integer
    For i = 1 To Size
        ws.Cells(startRow, i + 1).Value = i
        ws.Cells(startRow + i, 1).Value = i
    Next i
    With ws.Cells(startRow + 1, 2).Resize(Size, Size)
        .Value = mat
        .NumberFormat = numFormat
    End With
    ws.Cells(startRow, 1).Resize(Size + 1, Size + 1).HorizontalAlignment = xlCenter
    ws.Cells(startRow, 1).Resize(Size + 1, Size + 1).Borders.Weight = xlThin
End Sub

' ===========================================================================================
' PERFORMANCE WRAPPERS
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

