Attribute VB_Name = "mod_01_PickExp"
Option Explicit

' Module Name: PickExp
' Status: V61.0.25 - dictA2P stores both NumSubchans and PDU length
' Target: Excel 2024 LTSC
'
' Key changes from V61.0.20:
'   - dictA2P now stores Array(NtargetR umSubchans, PDU_Length)
'   - all dictA2P reads updated to use element(1) for PDU length
'   - HARQ-first + external CWLS flow preserved

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private data As Variant
Private idxTXID As Long
Private idxTXQ As Long
Private idxSFNCol As Long
Private idxLEN As Long
Private idxTXperSFN As Long
Private idxGen As Long
Private idxAvg As Long
Private idxTotLat As Long
Private idxGN As Long
Private idxRxCnt As Long
Private idxAppID As Long
Private txBitmap As String
Private bitmapLen As Long

Private rxStationIDs() As Long
Private rxDataColIdx() As Long
Private activeRxCount As Long
Private uniquePduSizes As Object

Private dictS2V As Object
Private dictVC As Object
Private dictA2P As Object
Private dictP2R As Object
Private dictP2Sigma As Object
Private sfnMap As Object
Private missingPduSizes As Object
Private nudgeLog() As Variant
Private nudgeCount As Long

Public dictTXStations As Object
Public dictRXStations As Object
Public runTX_SFN_CR As Boolean


Private Function MicroTimer() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency
    If QueryPerformanceFrequency(cyFreq) <> 0 Then
        Call QueryPerformanceCounter(cyTicks)
        If cyFreq > 0 Then
            MicroTimer = cyTicks / cyFreq
        End If
    End If
End Function

Private Sub EnsureSharedStationDicts()
    If dictTXStations Is Nothing Then Set dictTXStations = CreateObject("Scripting.Dictionary")
    If dictRXStations Is Nothing Then Set dictRXStations = CreateObject("Scripting.Dictionary")
End Sub

Sub PickExperimentFileAndMapData()
    Dim fd As Office.FileDialog
    Dim srcWB As Workbook, targetTable As ListObject
    Dim srcSheet As Worksheet, targetSheet As Worksheet
    Dim startRow As Long, endRow As Long, rowsToLoad As Long, i As Long, r As Long
    Dim srcData As Variant, idVal As Variant
    Dim c As Long
    Dim startTime As Double, totalProcTime As Double, pipelineStart As Double, pipelineTime As Double
    Dim iviVal As Double
    Dim genTime As Double
    
    Dim perfLog As Object
    Set perfLog = CreateObject("Scripting.Dictionary")
    Set missingPduSizes = CreateObject("Scripting.Dictionary")
    Set uniquePduSizes = CreateObject("Scripting.Dictionary")
    
    Dim uniqueVendors As Object
    Set uniqueVendors = CreateObject("Scripting.Dictionary")
    
    Dim harqDetectSeconds As Double
    Dim harqSplitSeconds As Double
    
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    
    Dim analysisChoice As String, msgMenu As String
Dim crChoice As VbMsgBoxResult

crChoice = MsgBox("Run TX_SFN Conflict Resolution?" & vbCrLf & _
                  "Default: Yes", vbYesNo + vbQuestion, "TX_SFN Conflict Resolution")
runTX_SFN_CR = (crChoice <> vbNo)

    msgMenu = "Select C-V2X Analysis routines to execute (e.g., 12345678):" & vbCrLf & _
          "1. GenerateLatencyAnalysis" & vbCrLf & _
          "2. GenerateLoadRxEfficiencyAnalysis" & vbCrLf & _
          "3. GenerateLoadTxEfficiencyAnalysis" & vbCrLf & _
          "4. GenerateRxGapPerAppAnalysis" & vbCrLf & _
          "5. GenerateTxGapPerAppAnalysis" & vbCrLf & _
          "6. GenerateTX_SFN_delta_histograms" & vbCrLf & _
          "7. GenerateRSSIMatrices" & vbCrLf & _
          "8. GenerateSpectralEfficiencyAnalysis"
    analysisChoice = InputBox(msgMenu, "C-V2X Routine Selection", "12345678")

    Set dictS2V = CreateObject("Scripting.Dictionary")
    Set dictVC = CreateObject("Scripting.Dictionary")
    Set dictA2P = CreateObject("Scripting.Dictionary")
    Set dictP2R = CreateObject("Scripting.Dictionary")
    Set dictP2Sigma = CreateObject("Scripting.Dictionary")
    Set sfnMap = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    txBitmap = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("tx_bitmap").Value
    bitmapLen = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("bitmap_len").Value
    On Error GoTo 0

    Dim gnPeriod As Double
    Dim gnFirstTX As Double
    Dim leapSecs As Double
    
    On Error Resume Next
    gnPeriod = Evaluate(ThisWorkbook.Names("GN_Time_32_bit_period__ms").RefersTo)
    leapSecs = Evaluate(ThisWorkbook.Names("Adj._for_leap_secs__ms").RefersTo)
    gnFirstTX = CDbl(ThisWorkbook.Names("GN_Time_of_First_TX__ms").RefersToRange.Value)
    On Error GoTo 0

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd.Show = False Then Exit Sub
    
    Set srcWB = Workbooks.Open(fd.SelectedItems(1), UpdateLinks:=0, ReadOnly:=True)
    Set srcSheet = srcWB.Sheets(1)
    
    Dim totalInFile As Long
    totalInFile = srcSheet.Cells(srcSheet.rows.count, "A").End(xlUp).Row - 1
    
    Dim rangeInput As String
    rangeInput = InputBox("Rows in file: " & Format(totalInFile, "#,###") & ". Range (e.g. 1-5000 or ALL):", "Range Selection", "ALL")
    
    Dim isParsed As Boolean
    isParsed = False
    Dim pts As Variant
    rangeInput = UCase$(Trim$(rangeInput))
    If rangeInput <> "ALL" And rangeInput <> "" Then
        pts = Split(rangeInput, "-")
        If UBound(pts) = 1 Then
            If IsNumeric(Trim$(pts(0))) And IsNumeric(Trim$(pts(1))) Then
                startRow = CLng(Trim$(pts(0))) + 1
                endRow = CLng(Trim$(pts(1))) + 1
                If startRow < 2 Then startRow = 2
                If endRow > totalInFile + 1 Then endRow = totalInFile + 1
                If endRow >= startRow Then isParsed = True
            End If
        End If
    End If
    If Not isParsed Then
        startRow = 2
        endRow = totalInFile + 1
    End If
    rowsToLoad = (endRow - startRow) + 1

    Set targetSheet = ThisWorkbook.Sheets("ExpResults")
    Set targetTable = targetSheet.ListObjects("ExpResultsTable")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    If Not targetTable.DataBodyRange Is Nothing Then
        targetTable.DataBodyRange.ClearContents
        If targetTable.ListRows.count > 1 Then
            targetTable.DataBodyRange.Offset(1, 0).Resize(targetTable.ListRows.count - 1).rows.Delete
        End If
    End If

    Dim anchorColIdx As Long
    On Error Resume Next
    anchorColIdx = targetTable.ListColumns("HARQ_indicator").Index + 1
    If Err.Number <> 0 Then
        Err.Clear
        anchorColIdx = targetTable.ListColumns("GN_TST").Index + 1
        If Err.Number <> 0 Then
            Err.Clear
            anchorColIdx = targetTable.ListColumns.count - 1
        End If
    End If
    On Error GoTo 0
    
    For c = targetTable.ListColumns.count To anchorColIdx Step -1
        targetTable.ListColumns(c).Delete
    Next c

    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    For i = 1 To srcSheet.Cells(1, srcSheet.Columns.count).End(xlToLeft).Column
        srcCols(UCase$(Trim$(srcSheet.Cells(1, i).Value))) = i
    Next i

    Dim foundIDs As String
    foundIDs = ""
    For i = 1 To srcSheet.Cells(1, srcSheet.Columns.count).End(xlToLeft).Column
        Dim hName As String
        hName = UCase$(srcSheet.Cells(1, i).Value)
        If Left$(hName, 6) = "RXTIME" Then
            foundIDs = foundIDs & Mid$(hName, 7) & ","
        End If
    Next i
    If Len(foundIDs) > 0 Then foundIDs = Left$(foundIDs, Len(foundIDs) - 1)

    Dim userIDs As String
    userIDs = InputBox("The following RX Station IDs were found: (" & foundIDs & "). Enter the IDs to process or ALL.", "RX Station Selection", "ALL")
    
    Dim selectedRXArray() As String
    If UCase$(Trim$(userIDs)) = "ALL" Then
        selectedRXArray = Split(foundIDs, ",")
    Else
        selectedRXArray = Split(userIDs, ",")
    End If

    Dim uniqueTXs As Object
    Set uniqueTXs = CreateObject("Scripting.Dictionary")
    
    Dim rawTXData As Variant
    rawTXData = srcSheet.Range(srcSheet.Cells(startRow, srcCols("TX_ID")), srcSheet.Cells(endRow, srcCols("TX_ID"))).Value
    
    For i = 1 To UBound(rawTXData, 1)
        If Not IsEmpty(rawTXData(i, 1)) Then
            uniqueTXs(CStr(rawTXData(i, 1))) = True
        End If
    Next i
    
    Dim sortedTXKeys() As String
    ReDim sortedTXKeys(0 To uniqueTXs.count - 1)
    
    Dim kIdx As Long
    kIdx = 0
    
    Dim vTxKey As Variant
    For Each vTxKey In uniqueTXs.Keys
        sortedTXKeys(kIdx) = CStr(vTxKey)
        kIdx = kIdx + 1
    Next vTxKey
    
    Dim pvt As String, tSwp As String, tLo As Long, tHi As Long
    Dim stkLo(1 To 64) As Long, stkHi(1 To 64) As Long, stkPtr As Long
    
    stkPtr = 1
    stkLo(1) = 0
    stkHi(1) = UBound(sortedTXKeys)
    
    Do While stkPtr > 0
        tLo = stkLo(stkPtr)
        tHi = stkHi(stkPtr)
        stkPtr = stkPtr - 1
        
        Do While tLo < tHi
            pvt = sortedTXKeys((tLo + tHi) \ 2)
            i = tLo
            r = tHi
            
            Do While i <= r
                If IsNumeric(sortedTXKeys(i)) And IsNumeric(pvt) Then
                    Do While CDbl(sortedTXKeys(i)) < CDbl(pvt)
                        i = i + 1
                    Loop
                    Do While CDbl(sortedTXKeys(r)) > CDbl(pvt)
                        r = r - 1
                    Loop
                Else
                    Do While sortedTXKeys(i) < pvt
                        i = i + 1
                    Loop
                    Do While sortedTXKeys(r) > pvt
                        r = r - 1
                    Loop
                End If
                
                If i <= r Then
                    tSwp = sortedTXKeys(i)
                    sortedTXKeys(i) = sortedTXKeys(r)
                    sortedTXKeys(r) = tSwp
                    i = i + 1
                    r = r - 1
                End If
            Loop
            
            If r - tLo > tHi - i Then
                If tLo < r Then
                    stkPtr = stkPtr + 1
                    stkLo(stkPtr) = tLo
                    stkHi(stkPtr) = r
                End If
                tLo = i
            Else
                If i < tHi Then
                    stkPtr = stkPtr + 1
                    stkLo(stkPtr) = i
                    stkHi(stkPtr) = tHi
                End If
                tHi = r
            End If
        Loop
    Loop
    
    Dim foundTXIDs As String
    foundTXIDs = Join(sortedTXKeys, ",")
    
    Dim userTXIDs As String
    userTXIDs = InputBox("The following TX Station IDs were found: (" & foundTXIDs & "). Enter the IDs to process or ALL.", "TX Station Selection", "ALL")
    
    Dim selectedTXArray() As String
    Dim filterTXDict As Object
    Set filterTXDict = CreateObject("Scripting.Dictionary")
    
    If UCase$(Trim$(userTXIDs)) = "ALL" Then
        selectedTXArray = sortedTXKeys
    Else
        selectedTXArray = Split(userTXIDs, ",")
    End If
    
    For Each idVal In selectedTXArray
        filterTXDict(Trim$(CStr(idVal))) = True
    Next idVal
    
    EnsureSharedStationDicts

    dictTXStations.RemoveAll
    dictRXStations.RemoveAll

    For Each idVal In selectedTXArray
        If Trim$(CStr(idVal)) <> "" Then
            dictTXStations(Trim$(CStr(idVal))) = True
        End If
    Next idVal

    For Each idVal In selectedRXArray
        If Trim$(CStr(idVal)) <> "" Then
            dictRXStations(Trim$(CStr(idVal))) = True
        End If
    Next idVal

    Dim rssiMissing As String
    rssiMissing = ""
    
    activeRxCount = 0
    ReDim rxStationIDs(1 To UBound(selectedRXArray) + 1)
    
    For Each idVal In selectedRXArray
        Dim curID As String
        curID = Trim$(idVal)
        If curID <> "" Then
            activeRxCount = activeRxCount + 1
            rxStationIDs(activeRxCount) = CLng(curID)
            targetTable.ListColumns.Add(anchorColIdx + activeRxCount - 1).Name = "RXTIME" & curID
        End If
    Next idVal

    For i = 1 To activeRxCount
        Dim curRSSIID As String
        curRSSIID = CStr(rxStationIDs(i))
        targetTable.ListColumns.Add(anchorColIdx + activeRxCount + i - 1).Name = "RSSI" & curRSSIID
        If Not srcCols.Exists("RSSI" & curRSSIID) Then
            rssiMissing = rssiMissing & curRSSIID & ", "
        End If
    Next i

    targetTable.ListColumns.Add.Name = "RX_COUNT"
    targetTable.ListColumns("RX_COUNT").Range.NumberFormat = "0"
    targetTable.ListColumns.Add.Name = "AVG_TOTAL_LATENCY"

    Dim stTable As ListObject
    Dim numericId As Long
    Set stTable = ThisWorkbook.Sheets("Exp Config & Data Proc Params").ListObjects("TX_RX_Station_IDs")
    
    If Not stTable.DataBodyRange Is Nothing Then
        stTable.ListColumns(2).DataBodyRange.ClearContents
        stTable.ListColumns(3).DataBodyRange.ClearContents
    End If
    
    For i = 0 To UBound(selectedTXArray)
        If IsNumeric(Trim$(selectedTXArray(i))) Then
            numericId = CLng(Trim$(selectedTXArray(i)))
            If numericId >= 1 And numericId <= stTable.ListRows.count Then
                stTable.DataBodyRange.Cells(numericId, 2).Value = numericId
            End If
        End If
    Next i
    
    For i = 1 To activeRxCount
        numericId = rxStationIDs(i)
        If numericId >= 1 And numericId <= stTable.ListRows.count Then
            stTable.DataBodyRange.Cells(numericId, 3).Value = numericId
        End If
    Next i
    
    ThisWorkbook.Names("Num_Rx_Stations").RefersTo = "=" & activeRxCount
    ThisWorkbook.Names("Num_Tx_Stations").RefersTo = "=" & (UBound(selectedTXArray) + 1)

    With targetTable
        idxAppID = .ListColumns("App_ID").Index
        idxTXID = .ListColumns("TX_ID").Index
        idxTXQ = .ListColumns("TXQTIME").Index
        idxSFNCol = .ListColumns("TX_SFN_est").Index
        idxLEN = .ListColumns("LEN").Index
        idxTXperSFN = .ListColumns("TXperSFN").Index
        idxGen = .ListColumns("MSG_GEN_TIME").Index
        idxAvg = .ListColumns("AVG TX1 RXTIMES").Index
        idxGN = .ListColumns("GN_TST").Index
        idxRxCnt = .ListColumns("RX_COUNT").Index
        idxTotLat = .ListColumns("AVG_TOTAL_LATENCY").Index
        
        ReDim rxDataColIdx(1 To activeRxCount)
        For i = 1 To activeRxCount
            rxDataColIdx(i) = .ListColumns("RXTIME" & rxStationIDs(i)).Index
        Next i
    End With

    LoadDictionariesFromThisWorkbook
    
    Dim mapK As Variant
    For Each mapK In dictS2V.Keys
        uniqueVendors(dictS2V(mapK)) = True
    Next mapK
    
    srcData = srcSheet.Cells(startRow, 1).Resize(rowsToLoad, srcSheet.Cells(1, srcSheet.Columns.count).End(xlToLeft).Column).Value
    
    Dim filteredCount As Long
    filteredCount = 0
    
    For r = 1 To rowsToLoad
        If filterTXDict.Exists(Trim$(CStr(srcData(r, srcCols("TX_ID"))))) Then
            filteredCount = filteredCount + 1
        End If
    Next r
    
    targetTable.Resize targetTable.HeaderRowRange.Resize(filteredCount + 1)
    data = targetTable.DataBodyRange.Value

    ReDim nudgeLog(1 To 50000, 1 To 4)
    nudgeCount = 0
    startTime = MicroTimer()

    Dim targetR As Long
    Dim txParamsChanged As Boolean
    Dim txLoopChanged As Boolean
    Dim continueLoop As Boolean
    Dim linRegCompleted As Boolean
    Dim linRegNeedsRerun As Boolean
    Dim currentTX As String
    Dim prelimRendered As Boolean

    continueLoop = True
    prelimRendered = False

    Do While continueLoop
        txLoopChanged = False
        linRegNeedsRerun = False

        ' -------------------------------
        ' 1) INITIAL SFN ESTIMATION PASS
        ' -------------------------------
        Set sfnMap = CreateObject("Scripting.Dictionary")
        For targetR = 1 To filteredCount
            GetSingleRowWLSCost targetR, 0
            ProcessInitialEstimation targetR
            AddToMap targetR, CLng(data(targetR, idxSFNCol))
        Next targetR

        ' ------------------------------------------------
        ' 2) PRELIMINARY RX/TX LATENCY OUTPUT FIRST
        ' ------------------------------------------------
        If Not prelimRendered Then
            Dim wsLogSheet As Worksheet
            On Error Resume Next
            Set wsLogSheet = ThisWorkbook.Sheets("TX_SFN est Log")
            If wsLogSheet Is Nothing Then
                Set wsLogSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("ExpResults"))
                wsLogSheet.Name = "TX_SFN est Log"
            End If
            On Error GoTo 0
            prelimRendered = True
        End If

        ' ============================================================
        ' IMPORTANT:
        ' Keep your existing preliminary plot/table rendering code HERE
        ' exactly as it already is, but do NOT call LinReg yet.
        ' ============================================================
    
     
    
    On Error Resume Next
    Set wsLogSheet = ThisWorkbook.Sheets("TX_SFN est Log")
    If wsLogSheet Is Nothing Then
        Set wsLogSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("ExpResults"))
        wsLogSheet.Name = "TX_SFN est Log"
    End If
    On Error GoTo 0
    
    Dim oldCht As ChartObject
    Dim pduKeys As Variant
    pduKeys = uniquePduSizes.Keys
    
    Dim pduIdx As Long
    Dim runningRowPos As Long
    Dim vendorKeys As Variant
    Dim vendorIdx As Long
    Dim finalChartRowBottom As Long
    Dim candidateRow As Long
    Dim renderWait As Double
    Dim vKey As Variant
    Dim loTxTable As ListObject
    Dim loRxTable As ListObject
    Dim parameterChanged As Boolean
    Dim anyParameterChanged As Boolean
    Dim rowWalk As Long
    Dim currentFilterPdu As Long
    
    Set loTxTable = ThisWorkbook.Sheets("Exp Config & Data Proc Params").ListObjects("VendorID2TXTproc")
    Set loRxTable = ThisWorkbook.Sheets("Exp Config & Data Proc Params").ListObjects("PDU2RXTprocVendorID")
    
    vendorKeys = uniqueVendors.Keys
    SortVariantLongArray pduKeys
    SortVariantStringArray vendorKeys
    
    anyParameterChanged = False
    
    Do
        wsLogSheet.Cells.Clear
        For Each oldCht In wsLogSheet.ChartObjects
            oldCht.Delete
        Next oldCht
        
        If anyParameterChanged Then
            wsLogSheet.Range("A1").Value = "PRE-WLS REPLOTTED LATENCY DISTRIBUTIONS (Updated Parameter Profiles)"
        Else
            wsLogSheet.Range("A1").Value = "PRE-WLS PRELIMINARY LATENCY DISTRIBUTIONS (Offset Verification Pass)"
        End If
        wsLogSheet.Range("A1").Font.Bold = True
        wsLogSheet.Range("A1").Font.Size = 14
        
        runningRowPos = 4
        finalChartRowBottom = runningRowPos + 4
        
        For vendorIdx = LBound(vendorKeys) To UBound(vendorKeys)
            If Trim$(CStr(vendorKeys(vendorIdx))) <> "" Then
                runningRowPos = RenderVendorPreWlsSection(wsLogSheet, data, rxDataColIdx, idxSFNCol, idxTXQ, idxTXID, idxLEN, _
                                                          dictS2V, Trim$(CStr(vendorKeys(vendorIdx))), pduKeys, runningRowPos)
            End If
        Next vendorIdx
        
        For Each oldCht In wsLogSheet.ChartObjects
            candidateRow = oldCht.BottomRightCell.Row + 2
            If candidateRow > finalChartRowBottom Then finalChartRowBottom = candidateRow
        Next oldCht
        If runningRowPos + 2 > finalChartRowBottom Then finalChartRowBottom = runningRowPos + 2
        
        wsLogSheet.Activate
        wsLogSheet.Range("A1").Select
        
        On Error Resume Next
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        On Error GoTo 0
        
        Application.ScreenUpdating = True
        DoEvents
        DoEvents
        
        renderWait = Timer + 1#
        Do While Timer < renderWait
            DoEvents
        Loop
        
        parameterChanged = False
        
        For Each vKey In uniqueVendors.Keys
            If Trim$(CStr(vKey)) <> "" Then
                If dictVC.Exists(CStr(vKey)) Then
                    Dim currentTxVals As Variant
                    currentTxVals = dictVC(CStr(vKey))
                    
                    Dim txPromptMsg As String
                    txPromptMsg = "Review plots rendered on the active 'TX_SFN est Log' sheet background." & vbCrLf & vbCrLf & _
                                  "Current TX parameters for Vendor " & vKey & ":" & vbCrLf & _
                                  "Mean (Tproc): " & currentTxVals(0) & " ms" & vbCrLf & _
                                  "Sigma: " & currentTxVals(1) & " ms" & vbCrLf & vbCrLf & _
                                  "To update, enter new values separated by a comma (e.g., 4.2,0.85)." & vbCrLf & _
                                  "Press Cancel or leave blank to retain values."
                    
                    Dim txUserResponse As String
                    txUserResponse = InputBox(txPromptMsg, "TX Parameter Modification - Vendor " & vKey, currentTxVals(0) & "," & currentTxVals(1))
                    
                    If Trim$(txUserResponse) <> "" And txUserResponse <> (currentTxVals(0) & "," & currentTxVals(1)) Then
                        Dim txSplit() As String
                        txSplit = Split(txUserResponse, ",")
                        If UBound(txSplit) = 1 Then
                            If IsNumeric(Trim$(txSplit(0))) And IsNumeric(Trim$(txSplit(1))) Then
                                Dim newTxMean As Double
                                Dim newTxSigma As Double
                                newTxMean = CDbl(Trim$(txSplit(0)))
                                newTxSigma = CDbl(Trim$(txSplit(1)))
                                
                                dictVC(CStr(vKey)) = Array(newTxMean, newTxSigma)
                                parameterChanged = True
                                anyParameterChanged = True
                                
                                For rowWalk = 1 To loTxTable.ListRows.count
                                    If Trim$(CStr(loTxTable.DataBodyRange.Cells(rowWalk, 1).Value)) = CStr(vKey) Then
                                        loTxTable.DataBodyRange.Cells(rowWalk, 2).Value = newTxMean
                                        loTxTable.DataBodyRange.Cells(rowWalk, 3).Value = newTxSigma
                                        Exit For
                                    End If
                                Next rowWalk
                            End If
                        End If
                    End If
                End If
            End If
        Next vKey
        
        For pduIdx = 0 To UBound(pduKeys)
            currentFilterPdu = CLng(pduKeys(pduIdx))
            For Each vKey In uniqueVendors.Keys
                If Trim$(CStr(vKey)) <> "" Then
                    Dim rKey As String
                    rKey = currentFilterPdu & "|" & vKey
                    
                    If dictP2R.Exists(rKey) Then
                        Dim currRxMean As Double
                        Dim currRxSig As Double
                        currRxMean = dictP2R(rKey)
                        currRxSig = IIf(dictP2Sigma.Exists(rKey), dictP2Sigma(rKey), 0#)
                        
                        Dim rxPromptMsg As String
                        rxPromptMsg = "Review plots rendered on the active 'TX_SFN est Log' sheet background." & vbCrLf & vbCrLf & _
                                      "Current RX parameters for PDU " & currentFilterPdu & "B (Vendor " & vKey & "):" & vbCrLf & _
                                      "Mean (Tproc): " & currRxMean & " ms" & vbCrLf & _
                                      "Sigma: " & currRxSig & " ms" & vbCrLf & vbCrLf & _
                                      "To update, enter new values separated by a comma (e.g., 3.1,0.45)." & vbCrLf & _
                                      "Press Cancel or leave blank to retain values."
                        
                        Dim rxUserResponse As String
                        rxUserResponse = InputBox(rxPromptMsg, "RX Parameter Modification - PDU: " & currentFilterPdu & "B, Vendor: " & vKey, currRxMean & "," & currRxSig)
                        
                        If Trim$(rxUserResponse) <> "" And rxUserResponse <> (currRxMean & "," & currRxSig) Then
                            Dim rxSplit() As String
                            rxSplit = Split(rxUserResponse, ",")
                            If UBound(rxSplit) = 1 Then
                                If IsNumeric(Trim$(rxSplit(0))) And IsNumeric(Trim$(rxSplit(1))) Then
                                    Dim newRxMean As Double
                                    Dim newRxSigma As Double
                                    newRxMean = CDbl(Trim$(rxSplit(0)))
                                    newRxSigma = CDbl(Trim$(rxSplit(1)))
                                    
                                    dictP2R(rKey) = newRxMean
                                    dictP2Sigma(rKey) = newRxSigma
                                    parameterChanged = True
                                    anyParameterChanged = True
                                    
                                    Dim targetMatrixColTime As Long
                                    Dim targetMatrixColSig As Long
                                    targetMatrixColTime = (CLng(vKey) * 2)
                                    targetMatrixColSig = targetMatrixColTime + 1
                                    
                                    For rowWalk = 1 To loRxTable.ListRows.count
                                        If CLng(loRxTable.DataBodyRange.Cells(rowWalk, 1).Value) = currentFilterPdu Then
                                            loRxTable.DataBodyRange.Cells(rowWalk, targetMatrixColTime).Value = newRxMean
                                            loRxTable.DataBodyRange.Cells(rowWalk, targetMatrixColSig).Value = newRxSigma
                                            Exit For
                                        End If
                                    Next rowWalk
                                End If
                            End If
                        End If
                    End If
                End If
            Next vKey
        Next pduIdx
        
        If parameterChanged Then
            Application.ScreenUpdating = False
            Set sfnMap = CreateObject("Scripting.Dictionary")
            For targetR = 1 To filteredCount
                GetSingleRowWLSCost targetR, 0
                ProcessInitialEstimation targetR
                AddToMap targetR, CLng(data(targetR, idxSFNCol))
            Next targetR
        End If
        
    Loop While parameterChanged
    
    totalProcTime = MicroTimer() - startTime
    
    
    For r = 1 To filteredCount
        Dim rxC As Long
        rxC = 0
        
        For i = 1 To activeRxCount
            If IsNumeric(data(r, rxDataColIdx(i))) Then
                If CDbl(data(r, rxDataColIdx(i))) <> 0 Then
                    rxC = rxC + 1
                End If
            End If
        Next i
        
        data(r, idxRxCnt) = rxC
    Next r
    
    targetTable.DataBodyRange.Value = data
    
    srcWB.Close False
    
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    Dim harqEnabled As Boolean
    Dim harqValue As Variant
    harqEnabled = False
    
    On Error Resume Next
    harqValue = ThisWorkbook.Names("HARQ").RefersToRange.Value
    If Err.Number = 0 Then
        If IsNumeric(harqValue) Then
            harqEnabled = (CLng(harqValue) = 1)
        End If
    End If
    Err.Clear
    
    If harqEnabled Then
        HARQDetection harqDetectSeconds
        If Err.Number = 0 Then perfLog("HARQDetection") = harqDetectSeconds
        Err.Clear
        
        HARQSplit harqSplitSeconds
        If Err.Number = 0 Then perfLog("HARQSplit") = harqSplitSeconds
        Err.Clear
    Else
        harqDetectSeconds = 0
        harqSplitSeconds = 0
    End If
    On Error GoTo 0
    
    data = targetTable.DataBodyRange.Value
    filteredCount = UBound(data, 1)
    
    Dim cwlsSeconds As Double
    cwlsSeconds = 0
    
    If runTX_SFN_CR Then
        TX_SFNConflictResolution data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, _
                               idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, _
                               activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, _
                               txBitmap, bitmapLen, cwlsSeconds
    End If

   ' 7. FINAL CALCULATIONS & WRITEBACK
 Application.StatusBar = "Writing mapping calculations back to RAM..."
 
 Dim sfnKey As Variant
 Dim prevFinalSfnKey As Variant
 Dim finalTxPerSFN As Long
 Dim rxSum As Double
 Dim rxCnt As Long
  
 For r = 1 To filteredCount
     sfnKey = data(r, idxSFNCol)
      
     If runTX_SFN_CR Then
         If r = 1 Then
             finalTxPerSFN = 1
         ElseIf IsNumeric(sfnKey) And IsNumeric(prevFinalSfnKey) And CLng(sfnKey) = CLng(prevFinalSfnKey) Then
             finalTxPerSFN = finalTxPerSFN + 1
         Else
             finalTxPerSFN = 1
         End If
         data(r, idxTXperSFN) = finalTxPerSFN
         prevFinalSfnKey = sfnKey
     Else
         If Not IsEmpty(sfnKey) And sfnMap.Exists(sfnKey) Then
             data(r, idxTXperSFN) = CLng(sfnMap(sfnKey).count)
         Else
             data(r, idxTXperSFN) = 0
         End If
     End If
      
     rxSum = 0
     rxCnt = 0
     
     For i = 1 To activeRxCount
         If IsNumeric(data(r, rxDataColIdx(i))) And data(r, rxDataColIdx(i)) <> 0 Then
             rxSum = rxSum + CDbl(data(r, rxDataColIdx(i)))
             rxCnt = rxCnt + 1
         End If
     Next i
     
     data(r, idxRxCnt) = rxCnt
     
     If rxCnt > 0 Then
         data(r, idxAvg) = rxSum / rxCnt
         data(r, idxTotLat) = (rxSum / rxCnt) - CDbl(data(r, idxGen))
     Else
         data(r, idxAvg) = vbNullString
         data(r, idxTotLat) = vbNullString
     End If
 Next r
 targetTable.Resize targetTable.HeaderRowRange.Resize(filteredCount + 1)
    targetTable.DataBodyRange.Value = data
    
    If Not targetTable.DataBodyRange Is Nothing Then
        filteredCount = targetTable.DataBodyRange.rows.count
        targetTable.Resize targetTable.HeaderRowRange.Resize(filteredCount + 1)
    End If
    
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    pipelineStart = MicroTimer()
    If analysisChoice <> "" Then
        RunPipeline analysisChoice, perfLog
    End If
    pipelineTime = MicroTimer() - pipelineStart
    
    Dim summaryMsg As String, k As Variant
    summaryMsg = "--- C-V2X MAPPING COMPLETE ---" & vbCrLf & _
                 "TX Stations Processed: " & (UBound(selectedTXArray) + 1) & vbCrLf & _
                 "RX Stations Processed: " & activeRxCount & vbCrLf & _
                 "HARQDetection Exec Time: " & Format(harqDetectSeconds, "0.000") & " s" & vbCrLf & _
                 "HARQSplit Exec Time: " & Format(harqSplitSeconds, "0.000") & " s" & vbCrLf & _
                 "TX_SFNConflictResolution Exec Time: " & Format(cwlsSeconds, "0.000") & " s" & vbCrLf & _
                 "Pre-HARQ Mapping Exec Time: " & Format(totalProcTime, "0.000") & " s" & vbCrLf & vbCrLf & _
                 "--- SUB-ROUTINE EXECUTION METRICS ---" & vbCrLf
                 
    If perfLog.count > 0 Then
        For Each k In perfLog.Keys
            summaryMsg = summaryMsg & k & ": " & Format(perfLog(k), "0.000") & " s" & vbCrLf
        Next k
    Else
        summaryMsg = summaryMsg & "(No downstream analysis modules selected)" & vbCrLf
    End If
    
    summaryMsg = summaryMsg & "---------------------------------------" & vbCrLf & _
                 "Total Analytics Pipeline Time: " & Format(pipelineTime, "0.000") & " seconds"
                 
    MsgBox summaryMsg, vbInformation, "Pipeline Performance Monitor"
End Sub

Private Function RenderVendorPreWlsSection(ws As Worksheet, dataBlock As Variant, rxCols() As Long, sfnIdx As Long, txqIdx As Long, txidIdx As Long, lenColIdx As Long, stToVenMap As Object, vendorID As String, pduKeys As Variant, startRowPos As Long) As Long
    vendorID = Trim$(CStr(vendorID))
    If vendorID = "" Then
        RenderVendorPreWlsSection = startRowPos
        Exit Function
    End If

    ws.Cells(startRowPos, 1).Value = "Vendor " & vendorID
    ws.Cells(startRowPos, 1).Font.Bold = True
    ws.Cells(startRowPos, 1).Font.Size = 14

    Dim rxMin As Double, rxMax As Double, rxStep As Double
    Dim txMin As Double, txMax As Double, txStep As Double
    On Error Resume Next
    rxMin = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("MIN_RX_MAC_latency").Value
    rxMax = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("MAX_RX_MAC_latency").Value
    rxStep = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("BIN_WIDTH_RX_MAC_latency").Value
    txMin = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("MIN_TX_MAC_latency").Value
    txMax = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("MAX_TX_MAC_latency").Value
    txStep = ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range("BIN_WIDTH_TX_MAC_latency").Value
    On Error GoTo 0

    If rxStep <= 0 Then rxStep = 1
    If txStep <= 0 Then txStep = 1
    If rxMax <= rxMin Then rxMax = rxMin + 100
    If txMax <= txMin Then txMax = txMin + 100

    Dim chartTopRow As Long: chartTopRow = startRowPos + 2
    Dim chartTopPos As Double: chartTopPos = ws.Cells(chartTopRow, 1).Top

    Dim chartWidth As Double: chartWidth = 360
    Dim chartHeight As Double: chartHeight = 260

    Dim rxChartLeft As Double: rxChartLeft = ws.Cells(chartTopRow, 1).Left
    Dim rxTableCol As Long: rxTableCol = 4
    Dim txChartLeft As Double: txChartLeft = ws.Cells(chartTopRow, 7).Left
    Dim txTableCol As Long: txTableCol = 10

    Dim rxTableLeft As Double: rxTableLeft = ws.Cells(chartTopRow, rxTableCol).Left
    Dim txTableLeft As Double: txTableLeft = ws.Cells(chartTopRow, txTableCol).Left

    Dim pduFilter As Long
    If IsArray(pduKeys) Then
        If UBound(pduKeys) >= LBound(pduKeys) Then
            pduFilter = CLng(pduKeys(LBound(pduKeys)))
        End If
    End If

    RenderSingleLatencyChart ws, dataBlock, rxCols, sfnIdx, txqIdx, txidIdx, lenColIdx, stToVenMap, vendorID, _
                             pduFilter, False, rxChartLeft, chartTopPos, chartWidth, chartHeight, _
                             rxMin, rxMax, rxStep

    RenderSingleLatencyChart ws, dataBlock, rxCols, sfnIdx, txqIdx, txidIdx, lenColIdx, stToVenMap, vendorID, _
                             0, True, txChartLeft, chartTopPos, chartWidth, chartHeight, _
                             txMin, txMax, txStep

    ws.Cells(chartTopRow, rxTableCol).Value = "RX MAC LATENCY SUMMARY - Vendor " & vendorID
    ws.Cells(chartTopRow, rxTableCol).Font.Bold = True
    ws.Cells(chartTopRow + 1, rxTableCol).Value = "Parameter"
    ws.Cells(chartTopRow + 1, rxTableCol + 1).Value = "Value"
    ws.Cells(chartTopRow + 1, rxTableCol).Font.Bold = True
    ws.Cells(chartTopRow + 1, rxTableCol + 1).Font.Bold = True
    WriteSingleLatencySummary ws, dataBlock, rxCols, sfnIdx, txqIdx, txidIdx, lenColIdx, stToVenMap, vendorID, _
                              pduFilter, False, chartTopRow + 2, rxTableCol

    ws.Cells(chartTopRow, txTableCol).Value = "TX MAC LATENCY SUMMARY - Vendor " & vendorID
    ws.Cells(chartTopRow, txTableCol).Font.Bold = True
    ws.Cells(chartTopRow + 1, txTableCol).Value = "Parameter"
    ws.Cells(chartTopRow + 1, txTableCol + 1).Value = "Value"
    ws.Cells(chartTopRow + 1, txTableCol).Font.Bold = True
    ws.Cells(chartTopRow + 1, txTableCol + 1).Font.Bold = True
    WriteSingleLatencySummary ws, dataBlock, rxCols, sfnIdx, txqIdx, txidIdx, lenColIdx, stToVenMap, vendorID, _
                              0, True, chartTopRow + 2, txTableCol

    RenderVendorPreWlsSection = chartTopRow + 14
End Function

Private Sub RenderSingleLatencyChart(ws As Worksheet, dataBlock As Variant, rxCols() As Long, sfnIdx As Long, txqIdx As Long, txidIdx As Long, lenColIdx As Long, stToVenMap As Object, vendorID As String, targetPduFilter As Long, isTXBlock As Boolean, chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double, bMin As Double, bMax As Double, bStep As Double)
    Dim lats() As Double
    Dim countVal As Long
    Dim n As Long, r As Long, i As Long
    Dim val As Double
    Dim mapValChart As Variant
    Dim rawLenVal As String
    Dim mappedPduStr As String
    
    ReDim lats(1 To UBound(dataBlock, 1) * (UBound(rxCols) + 1))
    countVal = 0
    
    For n = 1 To UBound(rxCols)
        If stToVenMap.Exists(CStr(n)) Then
            If stToVenMap(CStr(n)) = vendorID Then
                For r = 1 To UBound(dataBlock, 1)
                    Dim includeRow As Boolean
                    includeRow = False
                    
                    If isTXBlock Then
                        If CStr(dataBlock(r, txidIdx)) = CStr(n) Then
                            includeRow = True
                        End If
                    Else
                        rawLenVal = Trim$(CStr(dataBlock(r, lenColIdx)))
                        
                        If dictA2P.Exists(rawLenVal) Then
                            mapValChart = dictA2P(rawLenVal)
                            If IsArray(mapValChart) Then
                                mappedPduStr = Trim$(CStr(mapValChart(1)))
                            Else
                                mappedPduStr = Trim$(CStr(mapValChart))
                            End If
                        Else
                            mappedPduStr = rawLenVal
                        End If
                        
                        If IsNumeric(mappedPduStr) Then
                            If CLng(mappedPduStr) = targetPduFilter Then
                                includeRow = True
                            End If
                        End If
                    End If
                    
                    If includeRow Then
                        If isTXBlock Then
                            val = dataBlock(r, sfnIdx) - dataBlock(r, txqIdx)
                        Else
                            val = dataBlock(r, rxCols(n)) - dataBlock(r, sfnIdx)
                        End If
                        If val >= 0 Then
                            countVal = countVal + 1
                            lats(countVal) = val
                        End If
                    End If
                Next r
            End If
        End If
    Next n
    
    If countVal = 0 Then Exit Sub
    ReDim Preserve lats(1 To countVal)
    
    Dim xLabels() As Double, yFreq() As Double, yCDF() As Double, bCounts() As Long
    Dim nBins As Long
    nBins = CLng((bMax - bMin) / bStep) + 1
    If nBins > 2000 Then nBins = 2000
    ReDim xLabels(1 To nBins)
    ReDim yFreq(1 To nBins)
    ReDim yCDF(1 To nBins)
    ReDim bCounts(1 To nBins)
    
    For i = 1 To countVal
        Dim bIdx As Long
        bIdx = Int((lats(i) - bMin) / bStep) + 1
        If bIdx >= 1 And bIdx <= nBins Then bCounts(bIdx) = bCounts(bIdx) + 1
    Next i
    
    Dim curCum As Long
    curCum = 0
    For i = 1 To nBins
        xLabels(i) = bMin + (i - 1) * bStep
        yFreq(i) = bCounts(i)
        curCum = curCum + bCounts(i)
        yCDF(i) = curCum / countVal
    Next i
    
    Dim targetMu As Double: targetMu = 0
    Dim targetSigma As Double: targetSigma = 0
    
    If isTXBlock Then
        If dictVC.Exists(vendorID) Then
            Dim txParams As Variant
            txParams = dictVC(vendorID)
            targetMu = Abs(CDbl(txParams(0)))
            targetSigma = CDbl(txParams(1))
        End If
    Else
        Dim matrixKey As String
        matrixKey = targetPduFilter & "|" & vendorID
        If dictP2R.Exists(matrixKey) Then targetMu = dictP2R(matrixKey)
        If dictP2Sigma.Exists(matrixKey) Then targetSigma = dictP2Sigma(matrixKey)
    End If
    
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    
    With cht.Chart
        .HasTitle = True
        If isTXBlock Then
            .ChartTitle.Text = "Vendor " & vendorID & " PRELIMINARY TX MAC LATENCY"
        Else
            .ChartTitle.Text = "Vendor " & vendorID & " PRELIMINARY RX MAC LATENCY (PDU: " & targetPduFilter & "B)"
        End If
        
        With .SeriesCollection.NewSeries
            .Name = "Frequency"
            .Values = yFreq
            .XValues = xLabels
            .ChartType = xlColumnClustered
        End With
        .ChartGroups(1).GapWidth = 50
        
        With .SeriesCollection.NewSeries
            .Name = "CDF"
            .Values = yCDF
            .XValues = xLabels
            .ChartType = xlXYScatterLinesNoMarkers
            .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(0, 128, 0)
        End With
        
        Dim baseTprocSeriesIndex As Long
        Dim baseSigmaMinusSeriesIndex As Long
        Dim baseSigmaPlusSeriesIndex As Long
        baseTprocSeriesIndex = 0
        baseSigmaMinusSeriesIndex = 0
        baseSigmaPlusSeriesIndex = 0
        
        If targetMu > 0 Then
            With .SeriesCollection.NewSeries
                .Name = "Tproc"
                .Values = Array(0, 1)
                .XValues = Array(targetMu, targetMu)
                .ChartType = xlXYScatterLinesNoMarkers
                .AxisGroup = xlSecondary
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                .Format.Line.Weight = 1.5
            End With
            baseTprocSeriesIndex = .SeriesCollection.count
        End If
        
        If targetMu > 0 And targetSigma > 0 And (targetMu - targetSigma) >= bMin Then
            With .SeriesCollection.NewSeries
                .Name = "Tproc - Sigma"
                .Values = Array(0, 1)
                .XValues = Array(targetMu - targetSigma, targetMu - targetSigma)
                .ChartType = xlXYScatterLinesNoMarkers
                .AxisGroup = xlSecondary
                .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
                .Format.Line.DashStyle = msoLineDash
            End With
            baseSigmaMinusSeriesIndex = .SeriesCollection.count
        End If
        
        If targetMu > 0 And targetSigma > 0 And (targetMu + targetSigma) <= bMax Then
            With .SeriesCollection.NewSeries
                .Name = "Tproc + Sigma"
                .Values = Array(0, 1)
                .XValues = Array(targetMu + targetSigma, targetMu + targetSigma)
                .ChartType = xlXYScatterLinesNoMarkers
                .AxisGroup = xlSecondary
                .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
                .Format.Line.DashStyle = msoLineDash
            End With
            baseSigmaPlusSeriesIndex = .SeriesCollection.count
        End If
        
        If Not isTXBlock Then
            Dim chartHarqEnabled As Boolean
            Dim chartHarqValue As Variant
            chartHarqEnabled = False
            
            On Error Resume Next
            chartHarqValue = ThisWorkbook.Names("HARQ").RefersToRange.Value
            If Err.Number = 0 Then
                If IsNumeric(chartHarqValue) Then
                    chartHarqEnabled = (CLng(chartHarqValue) = 1)
                End If
            End If
            Err.Clear
            On Error GoTo 0
            
            If chartHarqEnabled And targetMu > 0 Then
                Dim harqOffset As Long
                Dim shiftedMu As Double
                Dim shiftedSigmaLow As Double
                Dim shiftedSigmaHigh As Double
                
                harqOffset = 1
                Do While (targetMu + harqOffset) <= bMax
                    shiftedMu = targetMu + harqOffset
                    
                    With .SeriesCollection.NewSeries
                        .Name = "Tproc"
                        .Values = Array(0, 1)
                        .XValues = Array(shiftedMu, shiftedMu)
                        .ChartType = xlXYScatterLinesNoMarkers
                        .AxisGroup = xlSecondary
                        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                        .Format.Line.Weight = 1.5
                    End With
                    
                    If targetSigma > 0 Then
                        shiftedSigmaLow = shiftedMu - targetSigma
                        shiftedSigmaHigh = shiftedMu + targetSigma
                        
                        If shiftedSigmaLow >= bMin Then
                            With .SeriesCollection.NewSeries
                                .Name = "Tproc - Sigma"
                                .Values = Array(0, 1)
                                .XValues = Array(shiftedSigmaLow, shiftedSigmaLow)
                                .ChartType = xlXYScatterLinesNoMarkers
                                .AxisGroup = xlSecondary
                                .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
                                .Format.Line.DashStyle = msoLineDash
                            End With
                        End If
                        
                        If shiftedSigmaHigh <= bMax Then
                            With .SeriesCollection.NewSeries
                                .Name = "Tproc + Sigma"
                                .Values = Array(0, 1)
                                .XValues = Array(shiftedSigmaHigh, shiftedSigmaHigh)
                                .ChartType = xlXYScatterLinesNoMarkers
                                .AxisGroup = xlSecondary
                                .Format.Line.ForeColor.RGB = RGB(0, 0, 255)
                                .Format.Line.DashStyle = msoLineDash
                            End With
                        End If
                    End If
                    
                    harqOffset = harqOffset + 1
                Loop
            End If
        End If
        
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Time (ms)"
            .TickLabels.Orientation = 90
        End With
        
        With .Axes(xlCategory, xlSecondary)
            .MinimumScale = bMin - (bStep / 2)
            .MaximumScale = (bMin + (nBins - 1) * bStep) + (bStep / 2)
            .TickLabelPosition = xlNone
            .Format.Line.Visible = msoFalse
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Frequency"
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "CDF Probability"
            .MinimumScale = 0
            .MaximumScale = 1
        End With
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        On Error Resume Next
        If .HasLegend Then
            Dim legendEntryCount As Long
            legendEntryCount = .Legend.LegendEntries.count
            For i = legendEntryCount To 1 Step -1
                If i <> 1 And i <> 2 Then
                    If i <> baseTprocSeriesIndex And i <> baseSigmaMinusSeriesIndex And i <> baseSigmaPlusSeriesIndex Then
                        .Legend.LegendEntries(i).Delete
                    End If
                End If
            Next i
        End If
        On Error GoTo 0
    End With
End Sub
