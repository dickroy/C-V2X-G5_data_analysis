Attribute VB_Name = "mod_11_TxEfficiencyLoad"
' ===========================================================================================
' Module: Load_Tx_Efficiency_Analysis
' Version: 6.2.2-INTEGRATED (COMPLETE MONOLITHIC BUILD)
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
Sub Run_GenerateLoadTxEfficiencyAnalysis()
    Dim startTime As Double: startTime = MicroTimer()
    
    ' Call core macro with pipeline parameters set to manual state
    Call GenerateLoadTxEfficiencyAnalysis(Nothing)
    
    ' Output execution timing only on manual runs
    Dim totalRunTime As Double: totalRunTime = MicroTimer() - startTime
    MsgBox "Tx Load & Efficiency Analysis Complete." & vbCrLf & _
           "Execution Time: " & Format(totalRunTime, "0.000") & " seconds", vbInformation, "Performance Monitor"
End Sub

' ===========================================================================================
' MAIN ENTRY POINT (Safe for automated pipelines calling GenerateLoadTxEfficiencyAnalysis(Nothing))
' ===========================================================================================
Sub GenerateLoadTxEfficiencyAnalysis(Optional ByRef logTable As Object = Nothing)
    Dim tStartExec As Double: tStartExec = MicroTimer() ' Start Performance Tracking
    
    Dim wsDest As Worksheet
    Dim wsApp As Worksheet
    Dim wsPdu As Worksheet
    Dim pduTable As ListObject
    Dim appTable As ListObject
    Dim tbsTable As Variant
    Dim pduArr As Variant
    Dim appArr As Variant
    Dim resultsArr() As Variant
    Dim stepRes() As Variant
    Dim aduSize As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim cho As ChartObject
    
    Dim lsMCS As Integer
    Dim lsNrb As Integer
    Dim lsTBS As Long
    Dim lsWaste As Double
    Dim hsMCS As Integer
    Dim hsNrb As Integer
    Dim hsTBS As Long
    Dim hsWaste As Double
    Const ndbps_p As Integer = 48
    Dim bits_total_p As Long
    Dim nsym_p As Long
    Dim pad_bytes_p As Double
    Dim tbs_bytes_p As Double
    Dim foundNrb As Integer
    
    ' Stat Variables (Primary: 10B+)
    Dim sumLS_Pct As Double, sumLS_Sq As Double, sumLS_PadB As Double, sumLS_PadB_Sq As Double
    Dim sumHS_Pct As Double, sumHS_Sq As Double, sumHS_PadB_Sq As Double, sumHS_PadB As Double
    Dim sum11p_Pct As Double, sum11p_Sq As Double, sum11p_PadB As Double, sum11p_PadB_Sq As Double
    Dim countLS As Long, countHS As Long, count11p As Long
    
    ' Toggle Excel engine optimizations on & push starting progress
    CoverScreen True
    UpdateProgressBar 5, "Initializing static TBS configurations..."
    
    tbsTable = GetStaticTBSTable()
    ReDim resultsArr(1 To 2491, 1 To 12)
    
    UpdateProgressBar 15, "Iterating ADU spectrum (10B to 2500B) for LS, HS, and DSRC..."
    For aduSize = 10 To 2500
        i = aduSize - 9
        resultsArr(i, 1) = aduSize
        
        ' --- 1. LOW SPEED (MCS 5-11, Max 2124B) ---
        If aduSize <= 2124 Then
            foundNrb = 999
            lsTBS = 999999
            For j = LBound(tbsTable, 1) To UBound(tbsTable, 1)
                If tbsTable(j, 1) >= 5 And tbsTable(j, 3) >= aduSize Then
                    If tbsTable(j, 2) < foundNrb Then
                        foundNrb = tbsTable(j, 2)
                    End If
                End If
            Next j
            
            If foundNrb <> 999 Then
                For j = LBound(tbsTable, 1) To UBound(tbsTable, 1)
                    If tbsTable(j, 1) >= 5 And tbsTable(j, 2) = foundNrb And tbsTable(j, 3) >= aduSize Then
                        If tbsTable(j, 3) < lsTBS Then
                            lsTBS = tbsTable(j, 3)
                            lsMCS = tbsTable(j, 1)
                            lsNrb = tbsTable(j, 2)
                        End If
                    End If
                Next j
                
                If lsMCS >= 11 Then
                    lsWaste = (lsTBS - aduSize) + (GetNPad(lsNrb) * 54)
                Else
                    lsWaste = (lsTBS - aduSize) + (GetNPad(lsNrb) * 27)
                End If
                
                resultsArr(i, 2) = lsMCS
                resultsArr(i, 3) = lsNrb
                resultsArr(i, 4) = lsTBS
                
                If lsMCS >= 11 Then
                    resultsArr(i, 5) = Application.WorksheetFunction.Max(0, (lsWaste / (lsTBS + (GetNPad(lsNrb) * 54))) + 0.01)
                Else
                    resultsArr(i, 5) = Application.WorksheetFunction.Max(0, (lsWaste / (lsTBS + (GetNPad(lsNrb) * 27))) + 0.01)
                End If
                
                If lsMCS < 11 Then
                    resultsArr(i, 11) = resultsArr(i, 5)
                Else
                    resultsArr(i, 12) = resultsArr(i, 5)
                End If
                
                sumLS_Pct = sumLS_Pct + resultsArr(i, 5)
                sumLS_Sq = sumLS_Sq + (resultsArr(i, 5) ^ 2)
                sumLS_PadB = sumLS_PadB + lsWaste
                sumLS_PadB_Sq = sumLS_PadB_Sq + (lsWaste ^ 2)
                countLS = countLS + 1
            End If
        End If

        ' --- 2. HIGH SPEED (MCS 0-7, Max 1479B) ---
        If aduSize <= 1479 Then
            foundNrb = 999
            hsTBS = 999999
            For j = LBound(tbsTable, 1) To UBound(tbsTable, 1)
                If tbsTable(j, 1) < 11 And tbsTable(j, 3) >= aduSize Then
                    If tbsTable(j, 2) < foundNrb Then
                        foundNrb = tbsTable(j, 2)
                    End If
                End If
            Next j
            
            If foundNrb <> 999 Then
                For j = LBound(tbsTable, 1) To UBound(tbsTable, 1)
                    If tbsTable(j, 1) < 11 And tbsTable(j, 2) = foundNrb And tbsTable(j, 3) >= aduSize Then
                        If tbsTable(j, 3) < hsTBS Then
                            hsTBS = tbsTable(j, 3)
                            hsMCS = tbsTable(j, 1)
                            hsNrb = tbsTable(j, 2)
                        End If
                    End If
                Next j
                
                hsWaste = (hsTBS - aduSize) + (GetNPad(hsNrb) * 27)
                resultsArr(i, 6) = hsMCS
                resultsArr(i, 7) = hsNrb
                resultsArr(i, 8) = hsTBS
                resultsArr(i, 9) = Application.WorksheetFunction.Max(0, (hsWaste / (hsTBS + (GetNPad(hsNrb) * 27))) - 0.01)
                
                sumHS_Pct = sumHS_Pct + resultsArr(i, 9)
                sumHS_Sq = sumHS_Sq + (resultsArr(i, 9) ^ 2)
                sumHS_PadB = sumHS_PadB + hsWaste
                sumHS_PadB_Sq = sumHS_PadB_Sq + (hsWaste ^ 2)
                countHS = countHS + 1
            End If
        End If

        ' --- 3. DSRC (Max 2500B) ---
        bits_total_p = 22 + (aduSize * 8)
        nsym_p = -Int(-bits_total_p / ndbps_p)
        pad_bytes_p = ((nsym_p * ndbps_p) - bits_total_p) / 8
        tbs_bytes_p = (nsym_p * ndbps_p) / 8
        resultsArr(i, 10) = Application.WorksheetFunction.Max(0, pad_bytes_p / tbs_bytes_p)
        
        sum11p_Pct = sum11p_Pct + resultsArr(i, 10)
        sum11p_Sq = sum11p_Sq + (resultsArr(i, 10) ^ 2)
        sum11p_PadB = sum11p_PadB + pad_bytes_p
        sum11p_PadB_Sq = sum11p_PadB_Sq + (pad_bytes_p ^ 2)
        count11p = count11p + 1
    Next aduSize
    
    UpdateProgressBar 40, "Setting up destination sheet layout and clearing historical charts..."
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Load_Tx Efficiency Analysis")
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsDest.Name = "Load_Tx Efficiency Analysis"
    End If
    
    wsDest.Cells.Clear
    For Each cho In wsDest.ChartObjects
        cho.Delete
    Next cho

    ' TABLE 1 (Cols A:D)
    wsDest.Range("A1").Value = "Stats (ADU >= 10B)"
    wsDest.Range("A2:D2").Value = Array("Metric", "C-V2X LS", "C-V2X HS", "DSRC")
    wsDest.Range("A3:A6").Value = Application.Transpose(Array("Mean Ineff %", "Std Dev Ineff %", "Mean Pad (B)", "Std Dev Pad (B)"))
    
    If countLS > 0 Then
        wsDest.Range("B3").Value = sumLS_Pct / countLS
        wsDest.Range("B4").Value = Sqr(Application.WorksheetFunction.Max(0, (sumLS_Sq / countLS) - ((sumLS_Pct / countLS) ^ 2)))
        wsDest.Range("B5").Value = sumLS_PadB / countLS
        wsDest.Range("B6").Value = Sqr(Application.WorksheetFunction.Max(0, (sumLS_PadB_Sq / countLS) - ((sumLS_PadB / countLS) ^ 2)))
    End If
    If countHS > 0 Then
        wsDest.Range("C3").Value = sumHS_Pct / countHS
        wsDest.Range("C4").Value = Sqr(Application.WorksheetFunction.Max(0, (sumHS_Sq / countHS) - ((sumHS_Pct / countHS) ^ 2)))
        wsDest.Range("C5").Value = sumHS_PadB / countHS
        wsDest.Range("C6").Value = Sqr(Application.WorksheetFunction.Max(0, (sumHS_PadB_Sq / countHS) - ((sumHS_PadB / countHS) ^ 2)))
    End If
    If count11p > 0 Then
        wsDest.Range("D3").Value = sum11p_Pct / count11p
        wsDest.Range("D4").Value = Sqr(Application.WorksheetFunction.Max(0, (sum11p_Sq / count11p) - ((sum11p_Pct / count11p) ^ 2)))
        wsDest.Range("D5").Value = sum11p_PadB / count11p
        wsDest.Range("D6").Value = Sqr(Application.WorksheetFunction.Max(0, (sum11p_PadB_Sq / count11p) - ((sum11p_PadB / count11p) ^ 2)))
    End If
    wsDest.Range("B3:D4").NumberFormat = "0.00%"
    wsDest.Range("B5:D6").NumberFormat = "0.00"

    ' Data Dump
    wsDest.Range("O1:Z1").Value = Array("ADU (B)", "LS MCS", "LS Nrb", "LS TBS", "LS Ineff %", "HS MCS", "HS Nrb", "HS TBS", "HS Ineff %", "DSRC Ineff %", "LS QPSK", "LS 16QAM")
    wsDest.Range("O2").Resize(2491, 12).Value = resultsArr
    wsDest.Range("S2:S2501,W2:W2501,X2:X2501,Y2:Y2501,Z2:Z2501").NumberFormat = "0.00%"
    wsDest.Columns("A:Z").AutoFit
    
    UpdateProgressBar 55, "Rendering Tx Inefficiency versus ADU scatter plots..."
    CreateEfficiencyScatterPlot wsDest, 140
    
    ' --- DEMAND & LOAD CALCULATION ---
    Dim wsSrc As Worksheet, targetTable As ListObject
    Dim dataArr As Variant, pduMapArr As Variant, nSchMap As Object
    Set wsSrc = ThisWorkbook.Sheets("ExpResults")
    Set wsApp = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
    Set wsPdu = ThisWorkbook.Sheets("PDU Size Table")
    Set targetTable = wsSrc.ListObjects("ExpResultsTable")
    Set pduTable = wsPdu.ListObjects("ADU2NumSubchansTable")
    Set appTable = wsApp.ListObjects("AppParams")
    
    Dim tStart As Double: tStart = wsApp.Range("T_start").Value
    Dim tStop As Double: tStop = wsApp.Range("T_stop").Value
    Dim tWin As Double: tWin = wsApp.Range("T_win_size").Value
    Dim tStep As Double: tStep = wsApp.Range("T_step").Value
    Dim nSchPerSubframe As Double: nSchPerSubframe = wsApp.Range("Nsch_per_subfr").Value
    Dim numSchSec As Double: numSchSec = wsApp.Range("Num_SCH_sec").Value
    
    dataArr = targetTable.DataBodyRange.Value
    pduMapArr = pduTable.DataBodyRange.Value
    Set nSchMap = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(pduMapArr, 1)
        nSchMap(CLng(pduMapArr(i, 1))) = CDbl(pduMapArr(i, 2))
    Next i
    
    Dim winCount As Long: winCount = Int((tStop - tStart) / tStep) + 1
    Dim loadResults(): ReDim loadResults(1 To winCount, 1 To 2)
    Dim colSFN As Long: colSFN = targetTable.ListColumns("TX_SFN_est").Index
    Dim colLEN As Long: colLEN = targetTable.ListColumns("LEN").Index
    Dim ptrStart As Long: ptrStart = 1
    Dim ptrEnd As Long: ptrEnd = 1
    
    For k = 1 To winCount
        Dim cL As Double: cL = tStart + (k - 1) * tStep
        Dim cU As Double: cU = cL + tWin
        If cU > tStop Then
            Exit For
        End If
        Dim sumNsch As Double: sumNsch = 0
        Do While ptrStart <= UBound(dataArr, 1)
            If dataArr(ptrStart, colSFN) >= cL Then
                Exit Do
            End If
            ptrStart = ptrStart + 1
        Loop
        If ptrEnd < ptrStart Then
            ptrEnd = ptrStart
        End If
        Do While ptrEnd <= UBound(dataArr, 1)
            If dataArr(ptrEnd, colSFN) > cU Then
                Exit Do
            End If
            ptrEnd = ptrEnd + 1
        Loop
        If ptrEnd > ptrStart Then
            For i = ptrStart To ptrEnd - 1
                If nSchMap.Exists(CLng(dataArr(i, colLEN))) Then
                    sumNsch = sumNsch + nSchMap(CLng(dataArr(i, colLEN)))
                End If
            Next i
        End If
        loadResults(k, 1) = cU / 1000
        loadResults(k, 2) = (sumNsch / (tWin * nSchPerSubframe))
    Next k
    
    ' Demand & PSR Calculation
    UpdateProgressBar 70, "Computing sliding application load demands..."
    Dim colADU As Long, colTTI As Long, colMID_App As Long, colResMID As Long, colResTime As Long, lc As ListColumn, cleanName As String
    For Each lc In appTable.ListColumns
        cleanName = UCase(Trim(lc.Name))
        If cleanName = "ADU SIZE (B)" Or cleanName = "ADU_SIZE" Then
            colADU = lc.Index
        ElseIf cleanName = "TTI(MS)" Or cleanName = "TTI" Then
            colTTI = lc.Index
        ElseIf cleanName = "APP_ID" Or cleanName = "MSG_ID" Then
            colMID_App = lc.Index
        End If
    Next lc
    
    colResMID = targetTable.ListColumns("App_ID").Index
    colResTime = targetTable.ListColumns("TXQTIME").Index
    appArr = appTable.DataBodyRange.Value
    Dim eCount As Long: eCount = 0
    Dim rawEvents(): ReDim rawEvents(1 To UBound(appArr, 1) * 2, 1 To 2)
    
    For i = 1 To UBound(appArr, 1)
        Dim curApp As Variant: curApp = appArr(i, colMID_App)
        Dim firstT As Double: firstT = 0
        Dim lastT As Double: lastT = 0
        For j = 1 To UBound(dataArr, 1)
            If dataArr(j, colResMID) = curApp Then
                If firstT = 0 Then
                    firstT = dataArr(j, colResTime)
                End If
                lastT = dataArr(j, colResTime)
            End If
        Next j
        Dim nSchs As Double: nSchs = 0
        If nSchMap.Exists(CLng(appArr(i, colADU))) Then
            nSchs = nSchMap(CLng(appArr(i, colADU)))
        End If
        Dim curDemand As Double: curDemand = (1000 * nSchs / appArr(i, colTTI)) / numSchSec
        eCount = eCount + 1
        rawEvents(eCount, 1) = firstT / 1000
        rawEvents(eCount, 2) = curDemand
        eCount = eCount + 1
        rawEvents(eCount, 1) = lastT / 1000
        rawEvents(eCount, 2) = -curDemand
    Next i
    
    ' Sort rawEvents by time
    For i = 1 To eCount - 1
        For j = i + 1 To eCount
            If rawEvents(i, 1) > rawEvents(j, 1) Then
                Dim tempT As Double: tempT = rawEvents(i, 1)
                Dim tempD As Double: tempD = rawEvents(i, 2)
                rawEvents(i, 1) = rawEvents(j, 1)
                rawEvents(i, 2) = rawEvents(j, 2)
                rawEvents(j, 1) = tempT
                rawEvents(j, 2) = tempD
            End If
        Next j
    Next i
    
    ' Generate Staircase Data for Demand Plot
    ReDim stepRes(1 To (eCount * 2), 1 To 2)
    Dim curSum As Double: curSum = 0
    Dim pPtr As Long: pPtr = 1
    For i = 1 To eCount
        stepRes(pPtr, 1) = rawEvents(i, 1)
        stepRes(pPtr, 2) = curSum
        pPtr = pPtr + 1
        curSum = curSum + rawEvents(i, 2)
        stepRes(pPtr, 1) = rawEvents(i, 1)
        stepRes(pPtr, 2) = curSum
        pPtr = pPtr + 1
    Next i
    
    ' Calculate PSR
    Dim psrResults(): ReDim psrResults(1 To winCount, 1 To 1)
    Dim currentD As Double: currentD = 0
    Dim evIdx As Long: evIdx = 1
    For k = 1 To winCount
        Do While evIdx <= eCount
            If rawEvents(evIdx, 1) <= loadResults(k, 1) Then
                currentD = currentD + rawEvents(evIdx, 2)
                evIdx = evIdx + 1
            Else
                Exit Do
            End If
        Loop
        If currentD > 0 Then
            psrResults(k, 1) = Application.WorksheetFunction.Min(1, loadResults(k, 2) / currentD)
        Else
            psrResults(k, 1) = 0
        End If
    Next k
    
    wsDest.Range("AB1:AF1").Value = Array("Time (Load)", "Load (%)", "Time (Demand)", "Demand (%)", "PSR (%)")
    wsDest.Range("AB2").Resize(winCount, 2).Value = loadResults
    wsDest.Range("AD2").Resize(UBound(stepRes, 1), 2).Value = stepRes
    wsDest.Range("AF2").Resize(winCount, 1).Value = psrResults
    wsDest.Range("AB2:AB" & winCount + 1).NumberFormat = "0.0"
    wsDest.Range("AC2:AC" & winCount + 1, "AE2:AF" & winCount + 1).NumberFormat = "0.00%"
    wsDest.Range("AD2:AD" & UBound(stepRes, 1) + 1).NumberFormat = "0.000"
    
    UpdateProgressBar 85, "Rendering Demand / Load plots & importing GPS synchronization logs..."
    CreateLoadDemandPlot wsDest, 660
    CreatePsrVsLoadPlot wsDest, 1180
    
    UpdateProgressBar 95, "Compiling localized subframe efficiency tables..."
    Call AddSecondaryStatsTable(wsDest, resultsArr)
    
    ' Restore regular Excel interface features
    CoverScreen False
    UpdateProgressBar 100, "Done!"
    
    ' Clear the status bar
    Application.StatusBar = False
    
    ' Return run duration if executed from a master macro pipeline
    Dim tElapsedTotal As Double: tElapsedTotal = MicroTimer() - tStartExec
    If Not logTable Is Nothing Then
        logTable("GenerateLoadTxEfficiencyAnalysis") = tElapsedTotal
    End If
End Sub

' ===========================================================================================
' PRIVATE HELPER SUBROUTINES AND FUNCTIONS
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

Private Sub AddSecondaryStatsTable(ByVal ws As Worksheet, ByRef dataArr() As Variant)
    ' Summation Variables
    Dim sumLS_Pct As Double, sumLS_Sq As Double, sumLS_Pad As Double, sumLS_PadSq As Double
    Dim sumHS_Pct As Double, sumHS_Sq As Double, sumHS_Pad As Double, sumHS_PadSq As Double
    Dim sum11p_Pct As Double, sum11p_Sq As Double, sum11p_Pad As Double, sum11p_PadSq As Double
    
    ' Individual Counters
    Dim countLS As Long, countHS As Long, count11p As Long
    
    Dim i As Long, adu As Long
    Dim lsPad As Double, hsPad As Double, dsrcPad As Double
    Dim bits_p As Long, nsym_p As Long
    
    For i = 1 To UBound(dataArr, 1)
        adu = dataArr(i, 1)
        If adu >= 100 Then
            ' LS Calculation (Max 2124B)
            If adu <= 2124 And dataArr(i, 4) > 0 Then
                countLS = countLS + 1
                If dataArr(i, 2) >= 11 Then
                    lsPad = (dataArr(i, 4) - adu) + (GetNPad(dataArr(i, 3)) * 54)
                Else
                    lsPad = (dataArr(i, 4) - adu) + (GetNPad(dataArr(i, 3)) * 27)
                End If
                sumLS_Pct = sumLS_Pct + dataArr(i, 5)
                sumLS_Sq = sumLS_Sq + (dataArr(i, 5) ^ 2)
                sumLS_Pad = sumLS_Pad + lsPad
                sumLS_PadSq = sumLS_PadSq + (lsPad ^ 2)
            End If
            
            ' HS Calculation (Max 1479B)
            If adu <= 1479 And dataArr(i, 8) > 0 Then
                countHS = countHS + 1
                hsPad = (dataArr(i, 8) - adu) + (GetNPad(dataArr(i, 7)) * 27)
                sumHS_Pct = sumHS_Pct + dataArr(i, 9)
                sumHS_Sq = sumHS_Sq + (dataArr(i, 9) ^ 2)
                sumHS_Pad = sumHS_Pad + hsPad
                sumHS_PadSq = sumHS_PadSq + (hsPad ^ 2)
            End If
            
            ' DSRC Calculation (Max 2500B)
            If adu <= 2500 Then
                count11p = count11p + 1
                bits_p = 22 + (adu * 8)
                nsym_p = -Int(-bits_p / 48)
                dsrcPad = ((nsym_p * 48) - bits_p) / 8
                sum11p_Pct = sum11p_Pct + dataArr(i, 10)
                sum11p_Sq = sum11p_Sq + (dataArr(i, 10) ^ 2)
                sum11p_Pad = sum11p_Pad + dsrcPad
                sum11p_PadSq = sum11p_PadSq + (dsrcPad ^ 2)
            End If
        End If
    Next i
    
    ws.Range("F1").Value = "Stats (ADU >= 100B)"
    ws.Range("F2:I2").Value = Array("Metric", "C-V2X LS", "C-V2X HS", "DSRC")
    ws.Range("F3:F6").Value = Application.Transpose(Array("Mean Ineff %", "Std Dev Ineff %", "Mean Pad (B)", "Std Dev Pad (B)"))
    
    ' Output LS
    If countLS > 0 Then
        ws.Range("G3").Value = sumLS_Pct / countLS
        ws.Range("G4").Value = Sqr(Application.WorksheetFunction.Max(0, (sumLS_Sq / countLS) - ((sumLS_Pct / countLS) ^ 2)))
        ws.Range("G5").Value = sumLS_Pad / countLS
        ws.Range("G6").Value = Sqr(Application.WorksheetFunction.Max(0, (sumLS_PadSq / countLS) - ((sumLS_Pad / countLS) ^ 2)))
    End If
    
    ' Output HS
    If countHS > 0 Then
        ws.Range("H3").Value = sumHS_Pct / countHS
        ws.Range("H4").Value = Sqr(Application.WorksheetFunction.Max(0, (sumHS_Sq / countHS) - ((sumHS_Pct / countHS) ^ 2)))
        ws.Range("H5").Value = sumHS_Pad / countHS
        ws.Range("H6").Value = Sqr(Application.WorksheetFunction.Max(0, (sumHS_PadSq / countHS) - ((sumHS_Pad / countHS) ^ 2)))
    End If
    
    ' Output DSRC
    If count11p > 0 Then
        ws.Range("I3").Value = sum11p_Pct / count11p
        ws.Range("I4").Value = Sqr(Application.WorksheetFunction.Max(0, (sum11p_Sq / count11p) - ((sum11p_Pct / count11p) ^ 2)))
        ws.Range("I5").Value = sum11p_Pad / count11p
        ws.Range("I6").Value = Sqr(Application.WorksheetFunction.Max(0, (sum11p_PadSq / count11p) - ((sum11p_Pad / count11p) ^ 2)))
    End If
    
    ws.Range("G3:I4").NumberFormat = "0.00%"
    ws.Range("G5:I6").NumberFormat = "0.00"
End Sub

Private Sub CreateLoadDemandPlot(ws As Worksheet, topPos As Double)
    Dim cho As ChartObject
    Set cho = ws.ChartObjects.Add(Left:=0, Width:=1000, Top:=topPos, Height:=500)
    Dim ser As Series
    Dim lastRLoad As Long
    lastRLoad = ws.Cells(ws.rows.count, "AB").End(xlUp).Row
    Dim lastRDemand As Long
    lastRDemand = ws.Cells(ws.rows.count, "AD").End(xlUp).Row
    
    Dim k As Long
    Dim psrVal As Double
    Dim r As Integer, g As Integer, b As Integer
    
    With cho.Chart
        .ChartType = 74 ' xlXYScatterLines
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' 1. Demand Series
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "Demand (%)"
            .XValues = ws.Range("AD2:AD" & lastRDemand)
            .Values = ws.Range("AE2:AE" & lastRDemand)
            .Format.Line.ForeColor.RGB = RGB(0, 112, 192)
            .Format.Line.Weight = 2
            .MarkerStyle = -4142
        End With
        
        ' 2. Load Series
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "Load (%)"
            .XValues = ws.Range("AB2:AB" & lastRLoad)
            .Values = ws.Range("AC2:AC" & lastRLoad)
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1.5
            .MarkerStyle = -4142
        End With
        
        ' 3. PSR Marker Series
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "PSR (%)"
            .XValues = ws.Range("AB2:AB" & lastRLoad)
            .Values = ws.Range("AF2:AF" & lastRLoad)
            .Format.Line.Visible = 0
            .MarkerStyle = 8
            .MarkerSize = 7
            
            For k = 1 To .Points.count
                psrVal = ws.Cells(k + 1, "AF").Value
                If psrVal >= 0.75 Then
                    r = Int(255 * (1 - (psrVal - 0.75) / 0.25))
                    g = 255
                    b = 0
                ElseIf psrVal > 0.5 Then
                    r = 255
                    g = Int(255 * (psrVal - 0.5) / 0.25)
                    b = 0
                Else
                    r = 255
                    g = 0
                    b = 0
                End If
                With .Points(k)
                    .MarkerBackgroundColor = RGB(r, g, b)
                    .MarkerForegroundColor = RGB(r, g, b)
                End With
            Next k
        End With
        
        ' ===================================================================================
        ' deGPS PLOTTING LOGIC (Using Column 3 with Distinct Line Colors & Legended Names)
        ' ===================================================================================
        Dim deGPSTable As ListObject
        Dim wsDesc As Worksheet
        Dim deGPSData As Variant
        Dim hasGPSData As Boolean: hasGPSData = False
        
        On Error Resume Next
        Set wsDesc = ThisWorkbook.Sheets("Test Description")
        If Not wsDesc Is Nothing Then
            Set deGPSTable = wsDesc.ListObjects("deGPSTable")
        End If
        
        ' Fallback sheet targets if "Test Description" sheet or table isn't found
        If deGPSTable Is Nothing Then Set deGPSTable = ThisWorkbook.Sheets("ExpResults").ListObjects("deGPSTable")
        If deGPSTable Is Nothing Then Set deGPSTable = ThisWorkbook.Sheets("Exp Config & Data Proc Params").ListObjects("deGPSTable")
        
        If Not deGPSTable Is Nothing Then
            If Not (deGPSTable.DataBodyRange Is Nothing) Then
                deGPSData = deGPSTable.DataBodyRange.Value
                hasGPSData = True
            End If
        End If
        On Error GoTo 0
        
        If hasGPSData Then
            Dim wsApp As Worksheet: Set wsApp = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
            Dim tStart As Double: tStart = wsApp.Range("T_start").Value
            Dim tStop As Double: tStop = wsApp.Range("T_stop").Value
            
            Dim gpsRow As Long
            Dim plotMinX As Double: plotMinX = tStart / 1000
            Dim plotMaxX As Double: plotMaxX = tStop / 1000
            
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
        .ChartTitle.Text = "Demand, Load, and PSR vs Time"
        .Axes(1).HasTitle = True
        .Axes(1).AxisTitle.Text = "Time (s)"
        .Axes(2).TickLabels.NumberFormat = "0%"
        .Axes(2).MaximumScale = 1
        
        ' Configure Legend at the Bottom to hold the Sync Loss items safely
        .HasLegend = True
        .Legend.Position = -4107
    End With
End Sub

Private Sub CreatePsrVsLoadPlot(ws As Worksheet, topPos As Double)
    Dim cho As ChartObject
    Set cho = ws.ChartObjects.Add(Left:=0, Width:=1000, Top:=topPos, Height:=500)
    Dim ser As Series
    Dim lastR As Long
    lastR = ws.Cells(ws.rows.count, "AB").End(xlUp).Row
    
    Dim k As Long
    Dim psrVal As Double
    Dim r As Integer, g As Integer, b As Integer
    
    With cho.Chart
        .ChartType = 75 ' xlXYScatter
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "PSR vs Load"
            .XValues = ws.Range("AC2:AC" & lastR)
            .Values = ws.Range("AF2:AF" & lastR)
            .Format.Line.Visible = 0
            .MarkerStyle = 8
            .MarkerSize = 7
            
            For k = 1 To .Points.count
                psrVal = ws.Cells(k + 1, "AF").Value
                If psrVal >= 0.75 Then
                    r = Int(255 * (1 - (psrVal - 0.75) / 0.25))
                    g = 255
                    b = 0
                ElseIf psrVal > 0.5 Then
                    r = 255
                    g = Int(255 * (psrVal - 0.5) / 0.25)
                    b = 0
                Else
                    r = 255
                    g = 0
                    b = 0
                End If
                With .Points(k)
                    .MarkerBackgroundColor = RGB(r, g, b)
                    .MarkerForegroundColor = RGB(r, g, b)
                End With
            Next k
        End With
        
        .HasTitle = True
        .ChartTitle.Text = "Packet Service Ratio vs Load"
        .Axes(1).HasTitle = True
        .Axes(1).AxisTitle.Text = "Load (%)"
        .Axes(1).TickLabels.NumberFormat = "0%"
        .Axes(2).HasTitle = True
        .Axes(2).AxisTitle.Text = "PSR (%)"
        .Axes(2).TickLabels.NumberFormat = "0%"
        .Axes(2).MaximumScale = 1
        .HasLegend = False
    End With
End Sub

Private Sub CreateEfficiencyScatterPlot(ws As Worksheet, topPos As Double)
    Dim cho As ChartObject
    Set cho = ws.ChartObjects.Add(Left:=0, Width:=1000, Top:=topPos, Height:=500)
    Dim ser As Series
    Dim j As Long
    Dim lastNrb As Integer
    Dim currentNrb As Integer
    Dim startAdu As Long
    Dim endAdu As Long
    
    With cho.Chart
        .ChartType = 75 ' xlXYScatter
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' 1. C-V2X High Speed (Limit 1479B)
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "C-V2X HS"
            .XValues = ws.Range("O2:O1471")
            .Values = ws.Range("W2:W1471")
            .MarkerStyle = 8
            .MarkerSize = 6
            .Format.Line.Visible = 0
            .MarkerBackgroundColor = RGB(128, 0, 128)
            .MarkerForegroundColor = RGB(128, 0, 128)
        End With
        
        ' 2. DSRC (Full range)
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "DSRC"
            .XValues = ws.Range("O2:O2501")
            .Values = ws.Range("X2:X2501")
            .MarkerStyle = 8
            .MarkerSize = 6
            .Format.Line.Visible = 0
            .MarkerBackgroundColor = RGB(0, 176, 80)
            .MarkerForegroundColor = RGB(0, 176, 80)
        End With
        
        ' 3. C-V2X Low Speed QPSK (Limit 2124B)
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "C-V2X LS QPSK"
            .XValues = ws.Range("O2:O2116")
            .Values = ws.Range("Y2:Y2116")
            .MarkerStyle = 8
            .MarkerSize = 6
            .Format.Line.Visible = 0
            .MarkerBackgroundColor = RGB(255, 0, 0)
            .MarkerForegroundColor = RGB(255, 0, 0)
        End With
        
        ' 4. C-V2X Low Speed 16-QAM (Limit 2124B)
        Set ser = .SeriesCollection.NewSeries
        With ser
            .Name = "C-V2X LS 16-QAM"
            .XValues = ws.Range("O2:O2116")
            .Values = ws.Range("Z2:Z2116")
            .MarkerStyle = 8
            .MarkerSize = 6
            .Format.Line.Visible = 0
            .MarkerBackgroundColor = RGB(255, 255, 0)
            .MarkerForegroundColor = RGB(255, 255, 0)
        End With
        
        .HasTitle = True
        .ChartTitle.Text = "Tx Inefficiency vs ADU Size"
        .Axes(1).MinimumScale = 0
        .Axes(1).MaximumScale = 2550
        .Axes(1).HasTitle = True
        .Axes(1).AxisTitle.Text = "ADU Size (B)"
        .Axes(2).TickLabels.NumberFormat = "0%"
        .Axes(2).MaximumScale = 1
        .HasLegend = True
        .Legend.Position = -4107
        
        DrawVerticalLine cho.Chart, 2124, True
        
        lastNrb = ws.Range("Q2").Value
        startAdu = ws.Range("O2").Value
        For j = 2 To 2116
            currentNrb = ws.Range("Q" & j).Value
            If (currentNrb <> lastNrb And currentNrb <> 0) Or j = 2116 Then
                endAdu = ws.Range("O" & j).Value
                If currentNrb <> lastNrb And endAdu < 2124 Then
                    DrawVerticalLine cho.Chart, endAdu, False
                End If
                If endAdu <= 2124 Then
                    AddRBLabel cho.Chart, (startAdu + endAdu) / 2, lastNrb
                End If
                startAdu = endAdu
                lastNrb = currentNrb
            End If
        Next j
    End With
End Sub

Private Sub DrawVerticalLine(ByVal cht As Chart, ByVal xVal As Double, ByVal isLimit As Boolean)
    Dim xPos As Double, yTop As Double, yBot As Double
    With cht
        xPos = .PlotArea.InsideLeft + ((xVal - .Axes(1).MinimumScale) / (.Axes(1).MaximumScale - .Axes(1).MinimumScale)) * .PlotArea.InsideWidth
        yTop = .PlotArea.InsideTop
        yBot = yTop + .PlotArea.InsideHeight
        With .Shapes.AddLine(xPos, yTop, xPos, yBot)
            If isLimit Then
                .Line.DashStyle = 1
                .Line.ForeColor.RGB = RGB(50, 50, 50)
                .Line.Weight = 2
            Else
                .Line.DashStyle = 2
                .Line.ForeColor.RGB = RGB(150, 150, 150)
                .Line.Weight = 1
            End If
        End With
    End With
End Sub

Private Sub AddRBLabel(ByVal cht As Chart, ByVal xMid As Double, ByVal nrbValue As Integer)
    Dim xPos As Double, yPos As Double, fullTxt As String, shp As Shape
    If nrbValue = 0 Then
        Exit Sub
    End If
    With cht
        xPos = .PlotArea.InsideLeft + ((xMid - .Axes(1).MinimumScale) / (.Axes(1).MaximumScale - .Axes(1).MinimumScale)) * .PlotArea.InsideWidth
        yPos = .PlotArea.InsideTop + 10
        fullTxt = "N_rb = " & nrbValue & vbCrLf & "(Low Speed)"
        Set shp = .Shapes.AddTextbox(1, xPos - 100, yPos, 200, 50)
        With shp.TextFrame2
            .TextRange.Text = fullTxt
            .TextRange.ParagraphFormat.Alignment = 2
            .VerticalAnchor = 1
            With .TextRange.Lines(1).Font
                .Size = 15
                .Bold = -1
            End With
            If .TextRange.Lines.count >= 2 Then
                With .TextRange.Lines(2).Font
                    .Size = 12
                    .Bold = -1
                End With
            End If
        End With
        shp.Fill.Visible = 0
        shp.Line.Visible = 0
    End With
End Sub

Private Function GetNPad(ByVal nrb As Integer) As Integer
    Select Case nrb
        Case 27
            GetNPad = 1
        Case 36
            GetNPad = 2
        Case 96
            GetNPad = 2
        Case Else
            GetNPad = 0
    End Select
End Function

Private Function GetStaticTBSTable() As Variant
    Dim arr(1 To 45, 1 To 3) As Variant, rC As Integer, mList As Variant, nrbs As Variant, tbsD As Variant, mV As Variant, nIdx As Integer
    mList = Array(0, 1, 2, 3, 4, 5, 6, 7, 11)
    nrbs = Array(18, 27, 36, 48, 96)
    tbsD = Array(Array(61, 93, 125, 165, 333), Array(79, 121, 161, 217, 437), Array(97, 149, 201, 269, 533), Array(129, 193, 261, 349, 693), Array(161, 241, 325, 437, 871), Array(193, 293, 389, 533, 1063), Array(233, 349, 469, 621, 1239), Array(277, 421, 549, 749, 1479), Null, Null, Null, Array(389, 597, 775, 1063, 2124))
    rC = 1
    For Each mV In mList
        For nIdx = 0 To 4
            arr(rC, 1) = CInt(mV)
            arr(rC, 2) = nrbs(nIdx)
            arr(rC, 3) = tbsD(CInt(mV))(nIdx)
            rC = rC + 1
        Next nIdx
    Next mV
    GetStaticTBSTable = arr
End Function

