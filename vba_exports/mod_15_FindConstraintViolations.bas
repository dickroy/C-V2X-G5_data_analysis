Attribute VB_Name = "mod_15_FindConstraintViolations"
Option Explicit

' Bookmark: FindConstraintViolations_v2.0.0
' Version: V2.0.0
' Status: in-memory validation:
'   1) TX_ID uniqueness within TX_SFN_est group
'   2) TX/RX overlap within TX_SFN_est group
'   3) Capacity per RX station
'   4) TX timing: TX_SFN_est >= ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement
'   5) RX timing: TX_SFN_est <= min(RXTIME) - RXTproc + 3*sigma
'
' Target: Excel 2024 LTSC

Public FCV_title As String
Public out_col As String

Private mFCV_CsetStartRows() As Long
Private mFCV_CsetEndRows() As Long
Private mFCV_CsetSFN() As Long
Private mFCV_GroupCount As Long
Private mFCV_CsetCount As Long
Private mFCV_WarningGroupCount As Long
Private mFCV_CacheValid As Boolean

Public Sub Run_FindConstraintViolations()
    Call FindConstraintViolations(0)
End Sub

Public Function FCV_HasCache() As Boolean
    FCV_HasCache = mFCV_CacheValid
End Function

Public Function FCV_GetGroupCount() As Long
    FCV_GetGroupCount = mFCV_GroupCount
End Function

Public Function FCV_GetCsetCount() As Long
    FCV_GetCsetCount = mFCV_CsetCount
End Function

Public Function FCV_GetWarningGroupCount() As Long
    FCV_GetWarningGroupCount = mFCV_WarningGroupCount
End Function

Public Function FCV_GetCsetStartRow(ByVal idx As Long) As Long
    If idx >= 1 And idx <= mFCV_CsetCount Then
        FCV_GetCsetStartRow = mFCV_CsetStartRows(idx)
    End If
End Function

Private Sub FCV_ResetCache()
    mFCV_GroupCount = 0
    mFCV_CsetCount = 0
    mFCV_WarningGroupCount = 0
    mFCV_CacheValid = False
    Erase mFCV_CsetStartRows
    Erase mFCV_CsetEndRows
    Erase mFCV_CsetSFN
End Sub

Private Sub FCV_AddCset(ByVal startRow As Long, ByVal endRow As Long, ByVal sfnVal As Long)
    mFCV_CsetCount = mFCV_CsetCount + 1
    ReDim Preserve mFCV_CsetStartRows(1 To mFCV_CsetCount)
    ReDim Preserve mFCV_CsetEndRows(1 To mFCV_CsetCount)
    ReDim Preserve mFCV_CsetSFN(1 To mFCV_CsetCount)
    mFCV_CsetStartRows(mFCV_CsetCount) = startRow
    mFCV_CsetEndRows(mFCV_CsetCount) = endRow
    mFCV_CsetSFN(mFCV_CsetCount) = sfnVal
End Sub

Public Function FindConstraintViolations(ByVal numViolations2Find As Long) As Long
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

    Dim numViolations As Long, writeRow As Long
    Dim NumConstraints As Long, NumWarnings As Long
    Dim groupHasViolation As Boolean
    Dim baseOutCol As Long, txTableCol As Long, rxTableColStart As Long
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

    On Error GoTo CleanFail
    NumConstraints = 3
    NumWarnings = 1

    Set wsExp = ThisWorkbook.Worksheets("ExpResults")
    Set tbl = wsExp.ListObjects("ExpResultsTable")

    Set wsCfg = ThisWorkbook.Worksheets("Exp Config & Data Proc Params")
    Set wsPdu = ThisWorkbook.Worksheets("PDU Size Table")
    Set pduTbl = wsPdu.ListObjects("ADU2NumSubchansTable")
    Set stationVendorTbl = wsCfg.ListObjects("StationID2VendorID")
    Set vendorTxtTbl = wsCfg.ListObjects("VendorID2TXTproc")
    Set pduRxtTbl = wsCfg.ListObjects("PDU2RXTprocVendorID")

    FCV_ResetCache

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("Conflict Resolution Log")
    On Error GoTo CleanFail
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsLog.Name = "Conflict Resolution Log"
    End If

    If LenB(Trim$(out_col)) = 0 Then out_col = "A"
    If LenB(Trim$(FCV_title)) = 0 Then FCV_title = "Group Analysis BEFORE Conflict Resolution"

    baseOutCol = wsLog.Range(UCase$(Trim$(out_col)) & "1").Column
    txTableCol = baseOutCol + 4
    rxTableColStart = baseOutCol + 8

    maxSch = GetWorkbookNameLong("Nsch_per_subfr")
    nRx = GetWorkbookNameLong("Num_Rx_Stations")
    maxTXDecrement = GetWorkbookNameDouble("maxTX_SFN_est_decrement")

    If tbl.DataBodyRange Is Nothing Or pduTbl.DataBodyRange Is Nothing Or _
       stationVendorTbl.DataBodyRange Is Nothing Or vendorTxtTbl.DataBodyRange Is Nothing Or _
       pduRxtTbl.DataBodyRange Is Nothing Then
        MsgBox "One or more required tables have no data rows.", vbExclamation, "Find Constraint Violations"
        FindConstraintViolations = 0
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

    ' LEN -> NumSubchans
    ' LEN -> PDU Length (B)
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

    Dim stationKey As Variant, vendorID As String
    For Each stationKey In dictStation2Vendor.Keys
        vendorID = CStr(dictStation2Vendor(stationKey))
        If dictVendor2TXT.Exists(vendorID) Then
            dictTXProc(CStr(stationKey)) = CDbl(dictVendor2TXT(vendorID))
        End If
    Next stationKey

    Dim stationIdNum As Long, vendorNum As Long, pduLenBytes As String
    Dim vendorProcCol As Long, vendorSigmaCol As Long, pduRow As Long
    Dim rxProcVal As Double, rxSigmaVal As Double

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

    If numViolations2Find = 0 Then
        wsLog.Range(wsLog.Cells(1, baseOutCol), wsLog.Cells(100000, baseOutCol + 25)).ClearContents
        wsLog.Cells(1, baseOutCol).Value = FCV_title
        wsLog.Cells(1, baseOutCol).Font.Bold = True
        wsLog.Cells(3, baseOutCol).Value = "Timestamp:"
        wsLog.Cells(3, baseOutCol + 1).Value = Now
        wsLog.Cells(4, baseOutCol).Value = "Issues Found:"
        wsLog.Cells(4, baseOutCol + 1).Value = 0
        wsLog.Cells(5, baseOutCol).Value = "Processing (s):"
        wsLog.Cells(5, baseOutCol + 1).Value = ""
        wsLog.Cells(6, baseOutCol).Value = "Num SFs with multiple TXs (Groups):"
        wsLog.Cells(6, baseOutCol + 1).Value = 0
        wsLog.Cells(7, baseOutCol).Value = "Num Groups with Constraint Violations (Catalog IDs 1-" & NumConstraints & "):"
        wsLog.Cells(7, baseOutCol + 1).Value = 0
        wsLog.Cells(8, baseOutCol).Value = "Warning Groups (Catalog IDs 1-" & NumWarnings & "):"
        wsLog.Cells(8, baseOutCol + 1).Value = 0
        wsLog.Cells(9, baseOutCol).Value = "Row"
        wsLog.Cells(9, baseOutCol + 1).Value = "SFN"
        wsLog.Cells(9, baseOutCol + 2).Value = "Type"
        wsLog.Cells(9, baseOutCol + 3).Value = "Description"
        wsLog.Range(wsLog.Cells(9, baseOutCol), wsLog.Cells(9, baseOutCol + 3)).Font.Bold = True
    End If

    wsLog.Range(wsLog.Cells(3, txTableCol), wsLog.Cells(100000, txTableCol + 1)).ClearContents
    wsLog.Range(wsLog.Cells(3, rxTableColStart), wsLog.Cells(100000, rxTableColStart + 30)).ClearContents

    wsLog.Cells(3, txTableCol).Value = "TX_ID"
    wsLog.Cells(3, txTableCol + 1).Value = "TXTproc"

    wsLog.Cells(3, rxTableColStart).Value = "Station_ID"
    wsLog.Cells(3, rxTableColStart + 1).Value = "PDU Length / RX Timing"

    writeRow = 4
    For Each stationKey In dictTXProc.Keys
        wsLog.Cells(writeRow, txTableCol).Value = stationKey
        wsLog.Cells(writeRow, txTableCol + 1).Value = dictTXProc(stationKey)
        writeRow = writeRow + 1
    Next stationKey

    ' RX table: Station_ID + one column per unique resolved PDU value
    Dim dictPDUUsed As Object, pduUsedKey As Variant
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

    wsLog.Cells(3, rxTableColStart).Value = "Station_ID"
    Dim pduHeaderCol As Long
    pduHeaderCol = rxTableColStart + 1
    For Each pduUsedKey In dictPDUUsed.Keys
        wsLog.Cells(3, pduHeaderCol).Value = "PDU=" & CStr(pduUsedKey)
        pduHeaderCol = pduHeaderCol + 1
    Next pduUsedKey

    Dim rxTableRow As Long, rxTableCol As Long, pduVal As String
    rxTableRow = 4
    For i = 1 To UBound(stationVendorData, 1)
        If Not IsEmpty(stationVendorData(i, 1)) Then
            stationIdNum = CLng(stationVendorData(i, 1))
            wsLog.Cells(rxTableRow, rxTableColStart).Value = stationIdNum

            rxTableCol = rxTableColStart + 1
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

    writeRow = 10
    numViolations = 0
    Application.StatusBar = "FindConstraintViolations V2.0.0: scanning ExpResultsTable..."

    i = 1
    Do While i <= UBound(data, 1)
        currentSFN = data(i, idxSFN)

        If IsEmpty(currentSFN) Or Trim$(CStr(currentSFN)) = "" Or val(currentSFN) = 0 Then
            i = i + 1
            GoTo NextGroup
        End If

        startRow = i
        Do While i < UBound(data, 1)
            If data(i + 1, idxSFN) <> currentSFN Then Exit Do
            i = i + 1
        Loop
        endRow = i
        mFCV_GroupCount = mFCV_GroupCount + 1
        groupHasViolation = False

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
            If IsNumeric(data(j, idxTXID)) Then
                txStationId = CLng(data(j, idxTXID))
                txCount = txCount + 1
                txIDs(txCount) = txStationId

                If IsNumeric(data(j, idxTXQTIME)) Then
                    txQTime = CDbl(data(j, idxTXQTIME))
                    txSfnVal = CDbl(currentSFN)

                    If dictTXProc.Exists(CStr(txStationId)) Then
                        txtProc = CDbl(dictTXProc(CStr(txStationId)))
                        If txSfnVal < (Round(txQTime + txtProc) - maxTXDecrement) Then
                            numViolations = numViolations + 1
                            groupHasViolation = True
                            If numViolations2Find = 0 Then
                                wsLog.Cells(writeRow, baseOutCol).Value = startRow
                                wsLog.Cells(writeRow, baseOutCol + 1).Value = currentSFN
                                wsLog.Cells(writeRow, baseOutCol + 2).Value = "TXTIME"
                                wsLog.Cells(writeRow, baseOutCol + 3).Value = "TX_SFN_est < ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement."
                                writeRow = writeRow + 1
                            Else
                                FindConstraintViolations = numViolations
                                Exit Function
                            End If
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
                    numViolations = numViolations + 1
                    groupHasViolation = True
                    If numViolations2Find = 0 Then
                        wsLog.Cells(writeRow, baseOutCol).Value = j
                        wsLog.Cells(writeRow, baseOutCol + 1).Value = currentSFN
                        wsLog.Cells(writeRow, baseOutCol + 2).Value = "RXTIME"
                        wsLog.Cells(writeRow, baseOutCol + 3).Value = "TX_SFN_est > min(RXTIME) - RXTproc + 3*sigma."
                        writeRow = writeRow + 1
                    Else
                        FindConstraintViolations = numViolations
                        Exit Function
                    End If
                End If
            End If
        Next j

        If Not IsUniqueLongList(txIDs, txCount) Then
            numViolations = numViolations + 1
            groupHasViolation = True
            If numViolations2Find = 0 Then
                wsLog.Cells(writeRow, baseOutCol).Value = startRow
                wsLog.Cells(writeRow, baseOutCol + 1).Value = currentSFN
                wsLog.Cells(writeRow, baseOutCol + 2).Value = 1
                wsLog.Cells(writeRow, baseOutCol + 3).Value = "TX_ID values are not unique within this TX_SFN_est group."
                writeRow = writeRow + 1
            Else
                FindConstraintViolations = numViolations
                Exit Function
            End If
        End If

        For st = 1 To nRx
            If rxHasAny(st) Then
                rxCount = rxCount + 1
                rxStationIDs(rxCount) = st
            End If
        Next st

        If HasOverlap(txIDs, txCount, rxStationIDs, rxCount) Then
            numViolations = numViolations + 1
            groupHasViolation = True
            If numViolations2Find = 0 Then
                wsLog.Cells(writeRow, baseOutCol).Value = startRow
                wsLog.Cells(writeRow, baseOutCol + 1).Value = currentSFN
                wsLog.Cells(writeRow, baseOutCol + 2).Value = 2
                wsLog.Cells(writeRow, baseOutCol + 3).Value = "Merged TX + RX station list is not unique within this TX_SFN_est group."
                writeRow = writeRow + 1
            Else
                FindConstraintViolations = numViolations
                Exit Function
            End If
        End If

        For st = 1 To nRx
            If capacitySum(st) > maxSch Then
                numViolations = numViolations + 1
                groupHasViolation = True
                If numViolations2Find = 0 Then
                    wsLog.Cells(writeRow, baseOutCol).Value = startRow
                    wsLog.Cells(writeRow, baseOutCol + 1).Value = currentSFN
                    wsLog.Cells(writeRow, baseOutCol + 2).Value = 3
                    wsLog.Cells(writeRow, baseOutCol + 3).Value = "St " & st & " sum " & capacitySum(st) & " > " & maxSch
                    writeRow = writeRow + 1
                Else
                    FindConstraintViolations = numViolations
                    Exit Function
                End If
            End If
        Next st

        If groupHasViolation Then FCV_AddCset startRow, endRow, CLng(currentSFN)

        i = i + 1
NextGroup:
    Loop

    If numViolations2Find = 0 Then
        wsLog.Cells(4, baseOutCol + 1).Value = numViolations
        wsLog.Cells(5, baseOutCol + 1).Value = Round(Timer - startTime, 3)
        wsLog.Cells(6, baseOutCol + 1).Value = mFCV_GroupCount
        wsLog.Cells(7, baseOutCol + 1).Value = mFCV_CsetCount
        wsLog.Cells(8, baseOutCol + 1).Value = mFCV_WarningGroupCount
        wsLog.Cells(1, baseOutCol).Font.Bold = True
        wsLog.Range(wsLog.Cells(1, baseOutCol), wsLog.Cells(100000, baseOutCol + 25)).Columns.AutoFit
        mFCV_CacheValid = True
    End If

    FindConstraintViolations = numViolations
    Exit Function

CleanFail:
    Application.StatusBar = False
    MsgBox "FindConstraintViolations failed:" & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbCritical, "Find Constraint Violations"
    FindConstraintViolations = -1
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
