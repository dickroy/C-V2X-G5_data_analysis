Attribute VB_Name = "mod_15_FindConstraintViolations"
Option Explicit

' Bookmark: FCV_v1_01
' Status: FCV v1.01 in-memory validation:
'   1) TX_ID uniqueness within TX_SFN_est group
'   2) TX/RX overlap within TX_SFN_est group
'   3) Capacity per RX station
'   4) TX timing: TX_SFN_est >= ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement
'   5) RX timing: TX_SFN_est <= min(RXTIME) - RXTproc + 3*sigma
'
' Severity model in v1.01:
'   - TX/RX uniqueness and capacity findings are violations.
'   - TXTIME and RXTIME findings are warnings.
'
' Target: Excel 2024 LTSC

Private Const FCV_LOG_SHEET As String = "TX_SFN Conflict Resolution Log"

Private mLastFCVViolationCount As Long
Private mLastFCVWarningCount As Long
Private mLastFCVGroupCount As Long

Public Sub Run_FCV()
    Call FCV(0, True)
End Sub

Public Sub Run_FindConstraintViolations()
    Call FCV(0, True)
End Sub

Public Function FCV(ByVal numViolations2Find As Long, Optional ByVal writeReport As Boolean = True) As Long
    Dim startTime As Double: startTime = Timer

    Dim wsExp As Worksheet, wsLog As Worksheet, wsCfg As Worksheet, wsPdu As Worksheet
    Dim tbl As ListObject, pduTbl As ListObject, stationVendorTbl As ListObject, vendorTxtTbl As ListObject, pduRxtTbl As ListObject
    Dim data As Variant, aduData As Variant, stationVendorData As Variant, vendorTxtData As Variant, pduRxtData As Variant

    Dim i As Long, j As Long, st As Long
    Dim startRow As Long, endRow As Long
    Dim currentSFN As Variant
    Dim maxSch As Long, nRx As Long, maxTXDecrement As Double
    Dim idxSFN As Long, idxTXID As Long, idxLEN As Long, idxTXQTIME As Long
    Dim rxColIdx() As Long

    Dim dictADU As Object          ' LEN -> NumSubchans
    Dim dictADU2PDU As Object      ' LEN -> PDU Length (B)
    Dim dictStation2Vendor As Object
    Dim dictVendor2TXT As Object
    Dim dictTXProc As Object
    Dim dictRXProc As Object
    Dim dictRXSigma As Object

    Dim numViolations As Long, numWarnings As Long, numGroups As Long
    Dim txIDs() As Long, rxStationIDs() As Long
    Dim txCount As Long, rxCount As Long
    Dim rxHasAny() As Boolean, capacitySum() As Long
    Dim rowLenNsch As Long, rowLenKey As String

    Dim txQTime As Double, txSfnVal As Double
    Dim txStationId As Long, txtProc As Double
    Dim rxMinTime As Double, rxMinSet As Boolean, rxTimeVal As Double
    Dim rxStationId As Long, rxProc As Double, rxSigma As Double
    Dim pduKey As String
    Dim rowRxMinStation As Long
    Dim rxThreshold As Double
    Dim stationKey As Variant, vendorID As String
    Dim stationIdNum As Long, vendorNum As Long, pduLenBytes As String
    Dim vendorProcCol As Long, vendorSigmaCol As Long, pduRow As Long
    Dim rxProcVal As Double, rxSigmaVal As Double

    Dim violationItems As Collection
    Dim warningItems As Collection

    On Error GoTo CleanFail

    mLastFCVViolationCount = 0
    mLastFCVWarningCount = 0
    mLastFCVGroupCount = 0

    Set wsExp = ThisWorkbook.Worksheets("ExpResults")
    Set tbl = wsExp.ListObjects("ExpResultsTable")

    Set wsCfg = ThisWorkbook.Worksheets("Exp Config & Data Proc Params")
    Set wsPdu = ThisWorkbook.Worksheets("PDU Size Table")
    Set pduTbl = wsPdu.ListObjects("ADU2NumSubchansTable")
    Set stationVendorTbl = wsCfg.ListObjects("StationID2VendorID")
    Set vendorTxtTbl = wsCfg.ListObjects("VendorID2TXTproc")
    Set pduRxtTbl = wsCfg.ListObjects("PDU2RXTprocVendorID")

    Set wsLog = GetFCVLogSheet()

    maxSch = GetWorkbookNameLong("Nsch_per_subfr")
    nRx = GetWorkbookNameLong("Num_Rx_Stations")
    maxTXDecrement = GetWorkbookNameDouble("maxTX_SFN_est_decrement")

    If tbl.DataBodyRange Is Nothing Or pduTbl.DataBodyRange Is Nothing Or _
       stationVendorTbl.DataBodyRange Is Nothing Or vendorTxtTbl.DataBodyRange Is Nothing Or _
       pduRxtTbl.DataBodyRange Is Nothing Then
        MsgBox "One or more required tables have no data rows.", vbExclamation, "FCV"
        FCV = 0
        Exit Function
    End If

    data = tbl.DataBodyRange.Value
    aduData = pduTbl.DataBodyRange.Value
    stationVendorData = stationVendorTbl.DataBodyRange.Value
    vendorTxtData = vendorTxtTbl.DataBodyRange.Value
    pduRxtData = pduRxtTbl.DataBodyRange.Value

    Set dictADU = CreateObject("Scripting.Dictionary")
    Set dictADU2PDU = CreateObject("Scripting.Dictionary")
    Set dictStation2Vendor = CreateObject("Scripting.Dictionary")
    Set dictVendor2TXT = CreateObject("Scripting.Dictionary")
    Set dictTXProc = CreateObject("Scripting.Dictionary")
    Set dictRXProc = CreateObject("Scripting.Dictionary")
    Set dictRXSigma = CreateObject("Scripting.Dictionary")

    dictADU.CompareMode = vbTextCompare
    dictADU2PDU.CompareMode = vbTextCompare
    dictStation2Vendor.CompareMode = vbTextCompare
    dictVendor2TXT.CompareMode = vbTextCompare
    dictTXProc.CompareMode = vbTextCompare
    dictRXProc.CompareMode = vbTextCompare
    dictRXSigma.CompareMode = vbTextCompare

    For i = 1 To UBound(aduData, 1)
        If Not IsEmpty(aduData(i, 1)) Then
            dictADU(CStr(aduData(i, 1))) = aduData(i, 2)
            dictADU2PDU(CStr(aduData(i, 1))) = aduData(i, 3)
        End If
    Next i

    For i = 1 To UBound(stationVendorData, 1)
        If Not IsEmpty(stationVendorData(i, 1)) Then
            dictStation2Vendor(CStr(stationVendorData(i, 1))) = CStr(stationVendorData(i, 2))
        End If
    Next i

    For i = 1 To UBound(vendorTxtData, 1)
        If Not IsEmpty(vendorTxtData(i, 1)) Then
            dictVendor2TXT(CStr(vendorTxtData(i, 1))) = CDbl(vendorTxtData(i, 2))
        End If
    Next i

    For Each stationKey In dictStation2Vendor.Keys
        vendorID = CStr(dictStation2Vendor(stationKey))
        If dictVendor2TXT.Exists(vendorID) Then
            dictTXProc(CStr(stationKey)) = CDbl(dictVendor2TXT(vendorID))
        End If
    Next stationKey

    For i = 1 To UBound(stationVendorData, 1)
        If Not IsEmpty(stationVendorData(i, 1)) Then
            stationIdNum = CLng(stationVendorData(i, 1))
            If IsNumeric(stationVendorData(i, 2)) Then
                vendorNum = CLng(stationVendorData(i, 2))
            Else
                vendorNum = 0
            End If

            If vendorNum >= 1 And vendorNum <= 3 Then
                vendorProcCol = 2 + (vendorNum - 1) * 2
                vendorSigmaCol = vendorProcCol + 1
            Else
                vendorProcCol = 0
                vendorSigmaCol = 0
            End If

            For pduRow = 1 To UBound(pduRxtData, 1)
                pduLenBytes = CStr(pduRxtData(pduRow, 1))
                If vendorProcCol > 0 Then
                    If IsNumeric(pduRxtData(pduRow, vendorProcCol)) Then
                        rxProcVal = CDbl(pduRxtData(pduRow, vendorProcCol))
                    Else
                        rxProcVal = 0#
                    End If
                    If IsNumeric(pduRxtData(pduRow, vendorSigmaCol)) Then
                        rxSigmaVal = CDbl(pduRxtData(pduRow, vendorSigmaCol))
                    Else
                        rxSigmaVal = 0#
                    End If
                Else
                    rxProcVal = 0#
                    rxSigmaVal = 0#
                End If

                dictRXProc(CStr(stationIdNum) & "|" & pduLenBytes) = rxProcVal
                dictRXSigma(CStr(stationIdNum) & "|" & pduLenBytes) = rxSigmaVal
            Next pduRow
        End If
    Next i

    idxSFN = tbl.ListColumns("TX_SFN_est").Index
    idxTXID = tbl.ListColumns("TX_ID").Index
    idxLEN = tbl.ListColumns("LEN").Index
    idxTXQTIME = tbl.ListColumns("TXQTIME").Index

    ReDim rxColIdx(1 To nRx)
    For i = 1 To nRx
        rxColIdx(i) = tbl.ListColumns("RXTIME" & CStr(i)).Index
    Next i

    If writeReport Then
        Set violationItems = New Collection
        Set warningItems = New Collection
        PrepareFCVReport wsLog, dictTXProc, dictRXProc, dictRXSigma, dictADU2PDU, data, idxLEN, stationVendorData
    End If

    numViolations = 0
    numWarnings = 0
    numGroups = 0
    Application.StatusBar = "FCV v1.01: scanning ExpResultsTable..."

    i = 1
    Do While i <= UBound(data, 1)
        currentSFN = data(i, idxSFN)

        If IsEmpty(currentSFN) Or Trim$(CStr(currentSFN)) = "" Then
            i = i + 1
            GoTo NextGroup
        End If

        numGroups = numGroups + 1
        startRow = i
        Do While i < UBound(data, 1)
            If data(i + 1, idxSFN) <> currentSFN Then Exit Do
            i = i + 1
        Loop
        endRow = i

        txCount = 0
        rxCount = 0
        ReDim txIDs(1 To (endRow - startRow + 1))
        ReDim rxStationIDs(1 To nRx)
        ReDim rxHasAny(1 To nRx)
        ReDim capacitySum(1 To nRx)

        For st = 1 To nRx
            rxHasAny(st) = False
            capacitySum(st) = 0
        Next st

        For j = startRow To endRow
            txSfnVal = CDbl(currentSFN)

            If IsNumeric(data(j, idxTXID)) Then
                txStationId = CLng(data(j, idxTXID))
                txCount = txCount + 1
                txIDs(txCount) = txStationId

                If IsNumeric(data(j, idxTXQTIME)) Then
                    txQTime = CDbl(data(j, idxTXQTIME))
                    If dictTXProc.Exists(CStr(txStationId)) Then
                        txtProc = CDbl(dictTXProc(CStr(txStationId)))
                        If txSfnVal < (Round(txQTime + txtProc) - maxTXDecrement) Then
                            numWarnings = numWarnings + 1
                            If writeReport Then warningItems.Add Array(startRow, currentSFN, "TXTIME", "TX_SFN_est < ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement.")
                        End If
                    End If
                End If
            End If

            rowLenNsch = 0
            If Not IsEmpty(data(j, idxLEN)) Then
                rowLenKey = CStr(data(j, idxLEN))
                If dictADU.Exists(rowLenKey) Then
                    If IsNumeric(dictADU(rowLenKey)) Then rowLenNsch = CLng(dictADU(rowLenKey))
                End If
            End If

            rxMinSet = False
            rxMinTime = 0#
            rowRxMinStation = 0

            For st = 1 To nRx
                If Not IsEmpty(data(j, rxColIdx(st))) And Trim$(CStr(data(j, rxColIdx(st)))) <> "" Then
                    rxHasAny(st) = True
                    capacitySum(st) = capacitySum(st) + rowLenNsch

                    If IsNumeric(data(j, rxColIdx(st))) Then
                        rxTimeVal = CDbl(data(j, rxColIdx(st)))
                        If Not rxMinSet Then
                            rxMinSet = True
                            rxMinTime = rxTimeVal
                            rowRxMinStation = st
                        ElseIf rxTimeVal < rxMinTime Then
                            rxMinTime = rxTimeVal
                            rowRxMinStation = st
                        End If
                    End If
                End If
            Next st

            If rxMinSet Then
                rowLenKey = CStr(data(j, idxLEN))
                If dictADU2PDU.Exists(rowLenKey) Then
                    pduKey = CStr(dictADU2PDU(rowLenKey))
                Else
                    pduKey = ""
                End If

                rxStationId = rowRxMinStation

                If dictRXProc.Exists(CStr(rxStationId) & "|" & pduKey) Then
                    rxProc = CDbl(dictRXProc(CStr(rxStationId) & "|" & pduKey))
                Else
                    rxProc = 0#
                End If

                If dictRXSigma.Exists(CStr(rxStationId) & "|" & pduKey) Then
                    rxSigma = CDbl(dictRXSigma(CStr(rxStationId) & "|" & pduKey))
                Else
                    rxSigma = 0#
                End If

                rxThreshold = rxMinTime - rxProc + 3# * rxSigma

                If txSfnVal > rxThreshold Then
                    numWarnings = numWarnings + 1
                    If writeReport Then warningItems.Add Array(j, currentSFN, "RXTIME", "TX_SFN_est > min(RXTIME) - RXTproc + 3*sigma.")
                End If
            End If
        Next j

        If Not IsUniqueLongList(txIDs, txCount) Then
            numViolations = numViolations + 1
            If writeReport Then violationItems.Add Array(startRow, currentSFN, "TX-TX", "TX_ID values are not unique within this TX_SFN_est group.")
            If numViolations2Find > 0 And numViolations >= numViolations2Find Then GoTo FinalizeFCV
        End If

        For st = 1 To nRx
            If rxHasAny(st) Then
                rxCount = rxCount + 1
                rxStationIDs(rxCount) = st
            End If
        Next st

        If HasOverlap(txIDs, txCount, rxStationIDs, rxCount) Then
            numViolations = numViolations + 1
            If writeReport Then violationItems.Add Array(startRow, currentSFN, "TX/RX", "Merged TX + RX station list is not unique within this TX_SFN_est group.")
            If numViolations2Find > 0 And numViolations >= numViolations2Find Then GoTo FinalizeFCV
        End If

        For st = 1 To nRx
            If capacitySum(st) > maxSch Then
                numViolations = numViolations + 1
                If writeReport Then violationItems.Add Array(startRow, currentSFN, "CAPACITY", "St " & st & " sum " & capacitySum(st) & " > " & maxSch)
                If numViolations2Find > 0 And numViolations >= numViolations2Find Then GoTo FinalizeFCV
            End If
        Next st

        i = i + 1
NextGroup:
    Loop

FinalizeFCV:
    mLastFCVViolationCount = numViolations
    mLastFCVWarningCount = numWarnings
    mLastFCVGroupCount = numGroups

    If writeReport Then
        WriteFCVSummary wsLog, numViolations, numWarnings, numGroups, Round(Timer - startTime, 3)
        WriteIssueCollection wsLog, 10, "VIOLATIONS", violationItems
        WriteIssueCollection wsLog, NextIssueSectionRow(10, violationItems), "WARNINGS", warningItems
        wsLog.Columns("I:Z").AutoFit
    End If

    Application.StatusBar = "FCV v1.01 done: " & numViolations & " violations, " & numWarnings & " warnings in " & Round(Timer - startTime, 2) & " s"
    FCV = numViolations
    Exit Function

CleanFail:
    Application.StatusBar = False
    MsgBox "FCV failed:" & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbCritical, "FCV"
    FCV = -1
End Function

Public Function FindConstraintViolations(ByVal numViolations2Find As Long) As Long
    FindConstraintViolations = FCV(numViolations2Find, (numViolations2Find = 0))
End Function

Public Function GetLastFCVViolationCount() As Long
    GetLastFCVViolationCount = mLastFCVViolationCount
End Function

Public Function GetLastFCVWarningCount() As Long
    GetLastFCVWarningCount = mLastFCVWarningCount
End Function

Public Function GetLastFCVGroupCount() As Long
    GetLastFCVGroupCount = mLastFCVGroupCount
End Function

Private Function GetFCVLogSheet() As Worksheet
    On Error Resume Next
    Set GetFCVLogSheet = ThisWorkbook.Worksheets(FCV_LOG_SHEET)
    On Error GoTo 0
    If GetFCVLogSheet Is Nothing Then
        Set GetFCVLogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetFCVLogSheet.Name = FCV_LOG_SHEET
    End If
End Function

Private Sub PrepareFCVReport(ByVal wsLog As Worksheet, ByVal dictTXProc As Object, ByVal dictRXProc As Object, ByVal dictRXSigma As Object, ByVal dictADU2PDU As Object, ByRef data As Variant, ByVal idxLEN As Long, ByRef stationVendorData As Variant)
    Dim writeRow As Long
    Dim stationKey As Variant
    Dim dictPDUUsed As Object, pduUsedKey As Variant
    Dim i As Long
    Dim rowLenKey As String
    Dim pduKey As String
    Dim pduHeaderCol As Long
    Dim rxTableRow As Long, rxTableCol As Long
    Dim stationIdNum As Long
    Dim pduVal As String
    Dim rxProcVal As Double
    Dim rxSigmaVal As Double

    wsLog.Range("I:Z").ClearContents
    wsLog.Range("M2:Z100000").ClearContents

    wsLog.Range("M2").Value = "TX_ID"
    wsLog.Range("N2").Value = "TXTproc"
    writeRow = 3
    For Each stationKey In dictTXProc.Keys
        wsLog.Cells(writeRow, "M").Value = stationKey
        wsLog.Cells(writeRow, "N").Value = dictTXProc(stationKey)
        writeRow = writeRow + 1
    Next stationKey

    Set dictPDUUsed = CreateObject("Scripting.Dictionary")
    dictPDUUsed.CompareMode = vbTextCompare
    For i = 1 To UBound(data, 1)
        If Not IsEmpty(data(i, idxLEN)) Then
            rowLenKey = CStr(data(i, idxLEN))
            If dictADU2PDU.Exists(rowLenKey) Then
                pduKey = CStr(dictADU2PDU(rowLenKey))
                dictPDUUsed(pduKey) = True
            End If
        End If
    Next i

    wsLog.Cells(2, "Q").Value = "Station_ID"
    pduHeaderCol = 18
    For Each pduUsedKey In dictPDUUsed.Keys
        wsLog.Cells(2, pduHeaderCol).Value = "PDU=" & CStr(pduUsedKey)
        pduHeaderCol = pduHeaderCol + 1
    Next pduUsedKey

    rxTableRow = 3
    For i = 1 To UBound(stationVendorData, 1)
        If Not IsEmpty(stationVendorData(i, 1)) Then
            stationIdNum = CLng(stationVendorData(i, 1))
            wsLog.Cells(rxTableRow, "Q").Value = stationIdNum
            rxTableCol = 18
            For Each pduUsedKey In dictPDUUsed.Keys
                pduVal = CStr(pduUsedKey)
                If dictRXProc.Exists(CStr(stationIdNum) & "|" & pduVal) Then
                    rxProcVal = CDbl(dictRXProc(CStr(stationIdNum) & "|" & pduVal))
                Else
                    rxProcVal = 0#
                End If
                If dictRXSigma.Exists(CStr(stationIdNum) & "|" & pduVal) Then
                    rxSigmaVal = CDbl(dictRXSigma(CStr(stationIdNum) & "|" & pduVal))
                Else
                    rxSigmaVal = 0#
                End If
                wsLog.Cells(rxTableRow, rxTableCol).Value = "RXTproc=" & rxProcVal & " sigma=" & rxSigmaVal
                rxTableCol = rxTableCol + 1
            Next pduUsedKey
            rxTableRow = rxTableRow + 1
        End If
    Next i
End Sub

Private Sub WriteFCVSummary(ByVal wsLog As Worksheet, ByVal numViolations As Long, ByVal numWarnings As Long, ByVal numGroups As Long, ByVal elapsedSeconds As Double)
    With wsLog.Range("I2")
        .Value = "FCV"
        .Font.Bold = True
        .Font.Size = 14
        .Offset(1, 0).Value = "Version:"
        .Offset(1, 1).Value = "v1.01"
        .Offset(2, 0).Value = "Timestamp:"
        .Offset(2, 1).Value = Now
        .Offset(3, 0).Value = "Violations Found:"
        .Offset(3, 1).Value = numViolations
        .Offset(4, 0).Value = "Warnings Found:"
        .Offset(4, 1).Value = numWarnings
        .Offset(5, 0).Value = "Groups Scanned:"
        .Offset(5, 1).Value = numGroups
        .Offset(6, 0).Value = "Processing (s):"
        .Offset(6, 1).Value = elapsedSeconds
    End With
End Sub

Private Sub WriteIssueCollection(ByVal wsLog As Worksheet, ByVal startRow As Long, ByVal sectionTitle As String, ByVal issues As Collection)
    Dim writeRow As Long
    Dim issueItem As Variant
    Dim idx As Long

    writeRow = startRow
    wsLog.Cells(writeRow, "I").Value = sectionTitle
    wsLog.Cells(writeRow, "I").Font.Bold = True
    writeRow = writeRow + 1

    wsLog.Cells(writeRow, "I").Value = "Row"
    wsLog.Cells(writeRow, "J").Value = "SFN"
    wsLog.Cells(writeRow, "K").Value = "Type"
    wsLog.Cells(writeRow, "L").Value = "Description"
    wsLog.Range(wsLog.Cells(writeRow, "I"), wsLog.Cells(writeRow, "L")).Font.Bold = True
    writeRow = writeRow + 1

    If issues Is Nothing Then
        wsLog.Cells(writeRow, "I").Value = "(none)"
        Exit Sub
    End If
    If issues.Count = 0 Then
        wsLog.Cells(writeRow, "I").Value = "(none)"
        Exit Sub
    End If

    For idx = 1 To issues.Count
        issueItem = issues(idx)
        wsLog.Cells(writeRow, "I").Value = issueItem(0)
        wsLog.Cells(writeRow, "J").Value = issueItem(1)
        wsLog.Cells(writeRow, "K").Value = issueItem(2)
        wsLog.Cells(writeRow, "L").Value = issueItem(3)
        writeRow = writeRow + 1
    Next idx
End Sub

Private Function NextIssueSectionRow(ByVal startRow As Long, ByVal issues As Collection) As Long
    NextIssueSectionRow = startRow + 3
    If Not issues Is Nothing Then NextIssueSectionRow = NextIssueSectionRow + issues.Count
End Function

Private Function GetWorkbookNameLong(ByVal nameText As String) As Long
    Dim nm As Name, expr As String, v As Variant
    On Error GoTo FailHard
    Set nm = ThisWorkbook.Names(nameText)
    expr = nm.RefersTo
    If Len(expr) > 0 Then If Left$(expr, 1) = "=" Then expr = Mid$(expr, 2)
    v = Application.Evaluate(expr)
    If IsError(v) Or Not IsNumeric(v) Then Err.Raise vbObjectError + 7100, "GetWorkbookNameLong", "Workbook name '" & nameText & "' did not evaluate to a numeric value."
    GetWorkbookNameLong = CLng(v)
    Exit Function
FailHard:
    Err.Raise vbObjectError + 7101, "GetWorkbookNameLong", "Could not resolve workbook name: " & nameText
End Function

Private Function GetWorkbookNameDouble(ByVal nameText As String) As Double
    Dim nm As Name, expr As String, v As Variant
    On Error GoTo FailHard
    Set nm = ThisWorkbook.Names(nameText)
    expr = nm.RefersTo
    If Len(expr) > 0 Then If Left$(expr, 1) = "=" Then expr = Mid$(expr, 2)
    v = Application.Evaluate(expr)
    If IsError(v) Or Not IsNumeric(v) Then Err.Raise vbObjectError + 7102, "GetWorkbookNameDouble", "Workbook name '" & nameText & "' did not evaluate to a numeric value."
    GetWorkbookNameDouble = CDbl(v)
    Exit Function
FailHard:
    Err.Raise vbObjectError + 7103, "GetWorkbookNameDouble", "Could not resolve workbook name: " & nameText
End Function

Private Function IsUniqueLongList(ByRef arr() As Long, ByVal n As Long) As Boolean
    Dim i As Long, j As Long
    If n <= 1 Then IsUniqueLongList = True: Exit Function
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i) = arr(j) Then IsUniqueLongList = False: Exit Function
        Next j
    Next i
    IsUniqueLongList = True
End Function

Private Function HasOverlap(ByRef txArr() As Long, ByVal txN As Long, ByRef rxArr() As Long, ByVal rxN As Long) As Boolean
    Dim i As Long, j As Long
    If txN = 0 Or rxN = 0 Then Exit Function
    For i = 1 To txN
        For j = 1 To rxN
            If txArr(i) = rxArr(j) Then HasOverlap = True: Exit Function
        Next j
    Next i
End Function
