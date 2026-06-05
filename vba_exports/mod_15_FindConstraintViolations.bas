Attribute VB_Name = "mod_15_FindConstraintViolations"
Option Explicit

' Bookmark: FindConstraintViolations_v21
' Status: in-memory validation:
'   1) TX_ID uniqueness within TX_SFN_est group
'   2) TX/RX overlap within TX_SFN_est group
'   3) Capacity per RX station
'   4) TX timing: TX_SFN_est >= ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement
'   5) RX timing: TX_SFN_est <= min(RXTIME) - RXTproc + 3*sigma
'
' Target: Excel 2024 LTSC

Public FCV_TotalGroups As Long
Public FCV_ViolatedGroups As Long
Public FCV_WarningGroups As Long
Public FCV_CsetCount As Long
Public FCV_CsetStarts() As Long
Public FCV_CsetEnds() As Long
Public FCV_CsetSFN() As Double
Public FCV_CsetCursor As Long
Public FCV_FilteredCount As Long
Public FCV_HasPrescan As Boolean

Public Sub Run_FindConstraintViolations()
    Call FindConstraintViolations(0)
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
    Dim violationWriteRow As Long, warningWriteRow As Long
    Dim warningRows() As Long, warningSFN() As Double, warningTypes() As String, warningDesc() As String
    Dim warningCount As Long, warningCapacity As Long
    Dim csetCapacity As Long
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
    Dim groupSize As Long
    Dim groupHasViolation As Boolean, groupHasWarning As Boolean
    Dim warningHeadingRow As Long

    On Error GoTo CleanFail

    Set wsExp = ThisWorkbook.Worksheets("ExpResults")
    Set tbl = wsExp.ListObjects("ExpResultsTable")

    Set wsCfg = ThisWorkbook.Worksheets("Exp Config & Data Proc Params")
    Set wsPdu = ThisWorkbook.Worksheets("PDU Size Table")
    Set pduTbl = wsPdu.ListObjects("ADU2NumSubchansTable")
    Set stationVendorTbl = wsCfg.ListObjects("StationID2VendorID")
    Set vendorTxtTbl = wsCfg.ListObjects("VendorID2TXTproc")
    Set pduRxtTbl = wsCfg.ListObjects("PDU2RXTprocVendorID")

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("TX_SFN est Log")
    On Error GoTo CleanFail
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsLog.Name = "TX_SFN est Log"
    End If

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

    FCV_TotalGroups = 0
    FCV_ViolatedGroups = 0
    FCV_WarningGroups = 0
    FCV_CsetCount = 0
    FCV_CsetCursor = 0
    FCV_FilteredCount = UBound(data, 1)
    FCV_HasPrescan = (numViolations2Find = 0)
    On Error Resume Next
    Erase FCV_CsetStarts
    Erase FCV_CsetEnds
    Erase FCV_CsetSFN
    On Error GoTo CleanFail

    If numViolations2Find = 0 Then
        wsLog.Range("I2:L100000").ClearContents
        With wsLog.Range("I2")
            .Value = "CONSTRAINT VIOLATIONS"
            .Font.Bold = True
            .Offset(1, 0).Value = "Timestamp:"
            .Offset(1, 1).Value = Now
            .Offset(2, 0).Value = "Num SFs with multiple TXs (Groups)"
            .Offset(2, 1).Value = 0
            .Offset(3, 0).Value = "Num of Groups with Constraint Violations"
            .Offset(3, 1).Value = 0
            .Offset(4, 0).Value = "Num of Groups with WARNINGS"
            .Offset(4, 1).Value = 0
            .Offset(5, 0).Value = "Processing time(s):"
            .Offset(5, 1).Value = ""
        End With
    End If

    ' TXTproc table 5 cols left from prior placement
    wsLog.Range("M2:N100000").ClearContents
    wsLog.Range("Q2:R100000").ClearContents

    wsLog.Range("M2").Value = "TX_ID"
    wsLog.Range("N2").Value = "TXTproc"

    wsLog.Range("Q2").Value = "Station_ID"
    wsLog.Range("R2").Value = "PDU Length / RX Timing"

    writeRow = 3
    For Each stationKey In dictTXProc.Keys
        wsLog.Cells(writeRow, "M").Value = stationKey
        wsLog.Cells(writeRow, "N").Value = dictTXProc(stationKey)
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

    wsLog.Cells(2, "Q").Value = "Station_ID"
    Dim pduHeaderCol As Long
    pduHeaderCol = 18 ' R
    For Each pduUsedKey In dictPDUUsed.Keys
        wsLog.Cells(2, pduHeaderCol).Value = "PDU=" & CStr(pduUsedKey)
        pduHeaderCol = pduHeaderCol + 1
    Next pduUsedKey

    Dim rxTableRow As Long, rxTableCol As Long, pduVal As String
    rxTableRow = 3
    For i = 1 To UBound(stationVendorData, 1)
        If Not IsEmpty(stationVendorData(i, 1)) Then
            stationIdNum = CLng(stationVendorData(i, 1))
            wsLog.Cells(rxTableRow, "Q").Value = stationIdNum

            rxTableCol = 18 ' R
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

    If numViolations2Find = 0 Then
        wsLog.Range("I9").Value = "VIOLATIONS"
        wsLog.Range("I9").Font.Bold = True
        wsLog.Range("I10").Value = "Row"
        wsLog.Range("J10").Value = "SFN"
        wsLog.Range("K10").Value = "Type"
        wsLog.Range("L10").Value = "Description"
        wsLog.Range("I10:L10").Font.Bold = True
    End If

    violationWriteRow = 11
    warningWriteRow = 0
    warningCount = 0
    warningCapacity = 0
    csetCapacity = 0
    numViolations = 0
    Application.StatusBar = "FindConstraintViolations v21: scanning ExpResultsTable..."

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
        groupSize = (endRow - startRow + 1)
        If groupSize <= 1 Then
            i = i + 1
            GoTo NextGroup
        End If

        FCV_TotalGroups = FCV_TotalGroups + 1
        groupHasViolation = False
        groupHasWarning = False

        txCount = 0
        rxCount = 0
        ReDim txIDs(1 To groupSize)
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
                                wsLog.Cells(violationWriteRow, "I").Value = startRow
                                wsLog.Cells(violationWriteRow, "J").Value = currentSFN
                                wsLog.Cells(violationWriteRow, "K").Value = "TXTIME"
                                wsLog.Cells(violationWriteRow, "L").Value = "TX_SFN_est < ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement."
                                violationWriteRow = violationWriteRow + 1
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
                    groupHasWarning = True
                    If numViolations2Find = 0 Then
                        warningCount = warningCount + 1
                        If warningCount = 1 Then
                            warningCapacity = 32
                            ReDim warningRows(1 To warningCapacity)
                            ReDim warningSFN(1 To warningCapacity)
                            ReDim warningTypes(1 To warningCapacity)
                            ReDim warningDesc(1 To warningCapacity)
                        ElseIf warningCount > warningCapacity Then
                            warningCapacity = warningCapacity * 2
                            ReDim Preserve warningRows(1 To warningCapacity)
                            ReDim Preserve warningSFN(1 To warningCapacity)
                            ReDim Preserve warningTypes(1 To warningCapacity)
                            ReDim Preserve warningDesc(1 To warningCapacity)
                        End If
                        warningRows(warningCount) = j
                        warningSFN(warningCount) = CDbl(currentSFN)
                        warningTypes(warningCount) = "RXTIME"
                        warningDesc(warningCount) = "TX_SFN_est > min(RXTIME) - RXTproc + 3*sigma."
                    End If
                End If
            End If
        Next j

        If Not IsUniqueLongList(txIDs, txCount) Then
            numViolations = numViolations + 1
            groupHasViolation = True
            If numViolations2Find = 0 Then
                wsLog.Cells(violationWriteRow, "I").Value = startRow
                wsLog.Cells(violationWriteRow, "J").Value = currentSFN
                wsLog.Cells(violationWriteRow, "K").Value = "TX-TX"
                wsLog.Cells(violationWriteRow, "L").Value = "TX_ID values are not unique within this TX_SFN_est group."
                violationWriteRow = violationWriteRow + 1
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
                wsLog.Cells(violationWriteRow, "I").Value = startRow
                wsLog.Cells(violationWriteRow, "J").Value = currentSFN
                wsLog.Cells(violationWriteRow, "K").Value = "TX/RX"
                wsLog.Cells(violationWriteRow, "L").Value = "Merged TX + RX station list is not unique within this TX_SFN_est group."
                violationWriteRow = violationWriteRow + 1
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
                    wsLog.Cells(violationWriteRow, "I").Value = startRow
                    wsLog.Cells(violationWriteRow, "J").Value = currentSFN
                    wsLog.Cells(violationWriteRow, "K").Value = "CAPACITY"
                    wsLog.Cells(violationWriteRow, "L").Value = "St " & st & " sum " & capacitySum(st) & " > " & maxSch
                    violationWriteRow = violationWriteRow + 1
                Else
                    FindConstraintViolations = numViolations
                    Exit Function
                End If
            End If
        Next st

        If groupHasViolation Then
            FCV_ViolatedGroups = FCV_ViolatedGroups + 1
            If FCV_CsetCount = 0 Then
                csetCapacity = 32
                ReDim FCV_CsetStarts(1 To csetCapacity)
                ReDim FCV_CsetEnds(1 To csetCapacity)
                ReDim FCV_CsetSFN(1 To csetCapacity)
            ElseIf (FCV_CsetCount + 1) > csetCapacity Then
                csetCapacity = csetCapacity * 2
                ReDim Preserve FCV_CsetStarts(1 To csetCapacity)
                ReDim Preserve FCV_CsetEnds(1 To csetCapacity)
                ReDim Preserve FCV_CsetSFN(1 To csetCapacity)
            End If
            FCV_CsetCount = FCV_CsetCount + 1
            FCV_CsetStarts(FCV_CsetCount) = startRow
            FCV_CsetEnds(FCV_CsetCount) = endRow
            FCV_CsetSFN(FCV_CsetCount) = CDbl(currentSFN)
        End If
        If groupHasWarning Then
            FCV_WarningGroups = FCV_WarningGroups + 1
        End If

        i = i + 1
NextGroup:
    Loop

    If numViolations2Find = 0 Then
        If FCV_CsetCount > 0 And csetCapacity > FCV_CsetCount Then
            ReDim Preserve FCV_CsetStarts(1 To FCV_CsetCount)
            ReDim Preserve FCV_CsetEnds(1 To FCV_CsetCount)
            ReDim Preserve FCV_CsetSFN(1 To FCV_CsetCount)
        End If

        If warningCount > 0 And warningCapacity > warningCount Then
            ReDim Preserve warningRows(1 To warningCount)
            ReDim Preserve warningSFN(1 To warningCount)
            ReDim Preserve warningTypes(1 To warningCount)
            ReDim Preserve warningDesc(1 To warningCount)
        End If

        wsLog.Range("J4").Value = FCV_TotalGroups
        wsLog.Range("J5").Value = FCV_ViolatedGroups
        wsLog.Range("J6").Value = FCV_WarningGroups
        wsLog.Range("J7").Value = Round(Timer - startTime, 3)

        warningHeadingRow = violationWriteRow + 1
        wsLog.Cells(warningHeadingRow, "I").Value = "WARNINGS"
        wsLog.Cells(warningHeadingRow, "I").Font.Bold = True
        wsLog.Cells(warningHeadingRow + 1, "I").Value = "Row"
        wsLog.Cells(warningHeadingRow + 1, "J").Value = "SFN"
        wsLog.Cells(warningHeadingRow + 1, "K").Value = "Type"
        wsLog.Cells(warningHeadingRow + 1, "L").Value = "Description"
        wsLog.Range("I" & (warningHeadingRow + 1) & ":L" & (warningHeadingRow + 1)).Font.Bold = True
        warningWriteRow = warningHeadingRow + 2
        For i = 1 To warningCount
            wsLog.Cells(warningWriteRow, "I").Value = warningRows(i)
            wsLog.Cells(warningWriteRow, "J").Value = warningSFN(i)
            wsLog.Cells(warningWriteRow, "K").Value = warningTypes(i)
            wsLog.Cells(warningWriteRow, "L").Value = warningDesc(i)
            warningWriteRow = warningWriteRow + 1
        End If

        wsLog.Columns("I:Z").AutoFit
    End If

    If numViolations2Find = 0 Then
        FindConstraintViolations = FCV_ViolatedGroups
    Else
        FindConstraintViolations = numViolations
    End If
    Exit Function

CleanFail:
    FCV_HasPrescan = False
    Application.StatusBar = False
    MsgBox "FindConstraintViolations failed:" & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbCritical, "Find Constraint Violations"
    FindConstraintViolations = -1
End Function

Public Sub FCV_ResetCsetCursor()
    FCV_CsetCursor = 0
End Sub

Public Function FCV_GetNextCsetGroup(ByRef startRow As Long, ByRef endRow As Long, ByRef groupSFN As Double) As Boolean
    FCV_GetNextCsetGroup = False
    If Not FCV_HasPrescan Then Exit Function
    If FCV_CsetCount <= 0 Then Exit Function
    If FCV_FilteredCount <= 0 Then Exit Function
    If FCV_CsetCursor >= FCV_CsetCount Then Exit Function

    FCV_CsetCursor = FCV_CsetCursor + 1
    startRow = FCV_CsetStarts(FCV_CsetCursor)
    endRow = FCV_CsetEnds(FCV_CsetCursor)
    groupSFN = FCV_CsetSFN(FCV_CsetCursor)
    FCV_GetNextCsetGroup = True
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
