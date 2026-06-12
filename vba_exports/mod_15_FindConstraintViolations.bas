Attribute VB_Name = "mod_15_FindConstraintViolations"
Option Explicit

' Bookmark: FindConstraintViolations
' Version: V2.0.1
' Status: in-memory validation:
'   1) TX_ID uniqueness within TX_SFN_est group         (type: TX-TX,      violation)
'   2) TX/RX overlap within TX_SFN_est group            (type: TX-RX,      violation)
'   3) Capacity per RX station                          (type: NSR-OL,     violation)
'   4) TX timing: TX_SFN_est >= ROUND(TXQTIME+TXTproc) (type: TXSFN-MTF,  warning)
'   5) RX timing: TX_SFN_est <= min(RXTIME)-RXTproc    (type: TXSFN>RXT,  warning)
'
' Target: Excel 2024 LTSC

Public FCV_title As String
Public out_col As String

Private mFCV_CsetStartRows() As Long
Private mFCV_CsetEndRows() As Long
Private mFCV_CsetSFN() As Long
Private mFCV_GroupCount As Long
Private mFCV_SFNCount As Long
Private mFCV_CsetCount As Long
Private mFCV_WarningGroupCount As Long
Private mFCV_CacheValid As Boolean

' ---------------------------------------------------------------------------
' Public entry points
' ---------------------------------------------------------------------------

Public Sub Run_FCV()
    ' Prompt user for BEFORE/AFTER placement in Conflict Resolution Log.
    ' AFTER is the default (second button / pressing Enter selects it).
    ' The only effect of this choice is the FCV_title / output placement.
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Is this FCV run BEFORE or AFTER Conflict Resolution is run on the raw data?" & vbCrLf & vbCrLf & _
                    "  [Yes] = BEFORE" & vbCrLf & _
                    "  [No]  = AFTER  (press Enter for default)", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "FCV Run Timing")
    out_col = "A"
    If answer = vbYes Then
        FCV_title = "Group Analysis BEFORE Conflict Resolution"
    Else
        FCV_title = "Group Analysis AFTER Conflict Resolution"
    End If
    Call FindConstraintViolations(0)
End Sub

Public Sub Run_FindConstraintViolations()
    Call FindConstraintViolations(0)
End Sub

' ---------------------------------------------------------------------------
' Cache accessors
' ---------------------------------------------------------------------------

Public Function FCV_HasCache() As Boolean
    FCV_HasCache = mFCV_CacheValid
End Function

Public Function FCV_GetSFNCount() As Long
    FCV_GetSFNCount = mFCV_SFNCount
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

' ---------------------------------------------------------------------------
' Cache management (private)
' ---------------------------------------------------------------------------

Private Sub FCV_ResetCache()
    mFCV_GroupCount = 0
    mFCV_SFNCount = 0
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

' ---------------------------------------------------------------------------
' Main function
' ---------------------------------------------------------------------------

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

    Dim numIssues As Long, numWarnings As Long, numViolations As Long
    Dim groupHasViolation As Boolean, groupHasWarning As Boolean
    Dim baseOutCol As Long
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

    Dim violationRows As Collection, warningRows As Collection

    On Error GoTo CleanFail

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
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = "Conflict Resolution Log"
    End If

    If LenB(Trim$(out_col)) = 0 Then out_col = "A"
    If LenB(Trim$(FCV_title)) = 0 Then FCV_title = "Group Analysis BEFORE Conflict Resolution"

    baseOutCol = wsLog.Range(UCase$(Trim$(out_col)) & "1").Column

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

        ' --- Title ---
        wsLog.Cells(1, baseOutCol).Value = FCV_title
        wsLog.Cells(1, baseOutCol).Font.Bold = True

        ' --- Summary counts (row 3-8, updated after scan) ---
        wsLog.Cells(3, baseOutCol).Value = "Timestamp:"
        wsLog.Cells(3, baseOutCol + 1).Value = Now
        wsLog.Cells(4, baseOutCol).Value = "Unique SFNs:"
        wsLog.Cells(4, baseOutCol + 1).Value = 0
        wsLog.Cells(5, baseOutCol).Value = "SFs with multiple TXs (Groups):"
        wsLog.Cells(5, baseOutCol + 1).Value = 0
        wsLog.Cells(6, baseOutCol).Value = "Groups with Constraint Violations:"
        wsLog.Cells(6, baseOutCol + 1).Value = 0
        wsLog.Cells(7, baseOutCol).Value = "Groups with WARNINGS:"
        wsLog.Cells(7, baseOutCol + 1).Value = 0
        wsLog.Cells(8, baseOutCol).Value = "Processing time (s):"
        wsLog.Cells(8, baseOutCol + 1).Value = ""

        ' --- Violation-type legend (rows 10-13) ---
        wsLog.Cells(10, baseOutCol).Value = "VIOLATION TYPES:"
        wsLog.Cells(10, baseOutCol).Font.Bold = True
        wsLog.Cells(11, baseOutCol).Value = "TX-TX"
        wsLog.Cells(11, baseOutCol + 1).Value = "Duplicate TX_ID values within SFN group"
        wsLog.Cells(12, baseOutCol).Value = "TX-RX"
        wsLog.Cells(12, baseOutCol + 1).Value = "TX/RX station ID overlap within SFN group"
        wsLog.Cells(13, baseOutCol).Value = "NSR-OL"
        wsLog.Cells(13, baseOutCol + 1).Value = "Number of subchannels exceeds Nsch capacity"

        ' --- Warning-type legend (rows 15-17) ---
        wsLog.Cells(15, baseOutCol).Value = "WARNING TYPES:"
        wsLog.Cells(15, baseOutCol).Font.Bold = True
        wsLog.Cells(16, baseOutCol).Value = "TXSFN>RXT"
        wsLog.Cells(16, baseOutCol + 1).Value = "TX_SFN_est > min(RXTIME) - RXTproc + 3*sigma"
        wsLog.Cells(17, baseOutCol).Value = "TXSFN-MTF"
        wsLog.Cells(17, baseOutCol + 1).Value = "TX_SFN_est < ROUND(TXQTIME + TXTproc) - maxTX_SFN_est_decrement"

        Set violationRows = New Collection
        Set warningRows = New Collection
    End If

    numIssues = 0
    numWarnings = 0
    numViolations = 0
    Application.StatusBar = "FindConstraintViolations V2.0.1: scanning ExpResultsTable..."

    i = 1
    Do While i <= UBound(data, 1)
        currentSFN = data(i, idxSFN)

        If IsEmpty(currentSFN) Or Trim$(CStr(currentSFN)) = "" Or Val(currentSFN) = 0 Then
            i = i + 1
            GoTo NextGroup
        End If

        startRow = i
        Do While i < UBound(data, 1)
            If data(i + 1, idxSFN) <> currentSFN Then Exit Do
            i = i + 1
        Loop
        endRow = i
        mFCV_SFNCount = mFCV_SFNCount + 1          ' total distinct SFN values seen
        If startRow < endRow Then mFCV_GroupCount = mFCV_GroupCount + 1  ' SFNs with >1 TX row
        groupHasViolation = False
        groupHasWarning = False

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
                            AppendFCVIssue violationRows, warningRows, startRow, currentSFN, "TXSFN-MTF", _
                                           numIssues, numWarnings, numViolations
                            groupHasWarning = True
                            If numViolations2Find <> 0 Then
                                FindConstraintViolations = numIssues
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
                    AppendFCVIssue violationRows, warningRows, j, currentSFN, "TXSFN>RXT", _
                                   numIssues, numWarnings, numViolations
                    groupHasWarning = True
                    If numViolations2Find <> 0 Then
                        FindConstraintViolations = numIssues
                        Exit Function
                    End If
                End If
            End If
        Next j

        If Not IsUniqueLongList(txIDs, txCount) Then
            AppendFCVIssue violationRows, warningRows, startRow, currentSFN, "TX-TX", _
                           numIssues, numWarnings, numViolations
            groupHasViolation = True
            If numViolations2Find <> 0 Then
                FindConstraintViolations = numIssues
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
            AppendFCVIssue violationRows, warningRows, startRow, currentSFN, "TX-RX", _
                           numIssues, numWarnings, numViolations
            groupHasViolation = True
            If numViolations2Find <> 0 Then
                FindConstraintViolations = numIssues
                Exit Function
            End If
        End If

        For st = 1 To nRx
            If capacitySum(st) > maxSch Then
                AppendFCVIssue violationRows, warningRows, startRow, currentSFN, "NSR-OL", _
                               numIssues, numWarnings, numViolations
                groupHasViolation = True
                If numViolations2Find <> 0 Then
                    FindConstraintViolations = numIssues
                    Exit Function
                End If
            End If
        Next st

        If groupHasViolation Then FCV_AddCset startRow, endRow, CLng(currentSFN)
        If groupHasWarning Then mFCV_WarningGroupCount = mFCV_WarningGroupCount + 1

        i = i + 1
NextGroup:
    Loop

    If numViolations2Find = 0 Then
        Dim writeRow As Long
        writeRow = WriteFCVIssueSection(wsLog, 19, "VIOLATIONS", violationRows, baseOutCol)
        WriteFCVIssueSection wsLog, writeRow, "WARNINGS", warningRows, baseOutCol

        wsLog.Cells(4, baseOutCol + 1).Value = mFCV_SFNCount
        wsLog.Cells(5, baseOutCol + 1).Value = mFCV_GroupCount
        wsLog.Cells(6, baseOutCol + 1).Value = mFCV_CsetCount
        wsLog.Cells(7, baseOutCol + 1).Value = mFCV_WarningGroupCount
        wsLog.Cells(8, baseOutCol + 1).Value = Round(Timer - startTime, 3)
        wsLog.Cells(1, baseOutCol).Font.Bold = True
        wsLog.Range(wsLog.Cells(1, baseOutCol), wsLog.Cells(100000, baseOutCol + 25)).Columns.AutoFit
        mFCV_CacheValid = True
    End If

    FindConstraintViolations = numIssues
    Exit Function

CleanFail:
    Application.StatusBar = False
    MsgBox "FindConstraintViolations failed:" & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbCritical, "Find Constraint Violations"
    FindConstraintViolations = -1
End Function

' ---------------------------------------------------------------------------
' Issue collection helper
' ---------------------------------------------------------------------------

Private Sub AppendFCVIssue(ByRef violationRows As Collection, ByRef warningRows As Collection, _
                           ByVal rowValue As Variant, ByVal sfnValue As Variant, _
                           ByVal issueType As String, _
                           ByRef numIssues As Long, ByRef numWarnings As Long, ByRef numViolations As Long)
    Dim issueEntry As Variant
    issueEntry = Array(rowValue, sfnValue, issueType)
    numIssues = numIssues + 1

    If IsFCVWarning(issueType) Then
        numWarnings = numWarnings + 1
        If Not warningRows Is Nothing Then warningRows.Add issueEntry
    Else
        numViolations = numViolations + 1
        If Not violationRows Is Nothing Then violationRows.Add issueEntry
    End If
End Sub

Private Function IsFCVWarning(ByVal issueType As String) As Boolean
    ' TXSFN>RXT and TXSFN-MTF are timing warnings; TX-TX, TX-RX, NSR-OL are hard violations.
    IsFCVWarning = (StrComp(Trim$(issueType), "TXSFN>RXT", vbTextCompare) = 0 Or _
                    StrComp(Trim$(issueType), "TXSFN-MTF", vbTextCompare) = 0)
End Function

Private Function WriteFCVIssueSection(ByVal wsLog As Worksheet, ByVal startRow As Long, _
                                      ByVal sectionTitle As String, ByVal issues As Collection, _
                                      ByVal baseCol As Long) As Long
    Dim r As Long, i As Long
    Dim issueEntry As Variant

    If issues Is Nothing Or issues.Count = 0 Then
        WriteFCVIssueSection = startRow
        Exit Function
    End If

    r = startRow
    wsLog.Cells(r, baseCol).Value = sectionTitle
    wsLog.Cells(r, baseCol).Font.Bold = True
    r = r + 1

    wsLog.Cells(r, baseCol).Value = "Row"
    wsLog.Cells(r, baseCol + 1).Value = "SFN"
    wsLog.Cells(r, baseCol + 2).Value = "Type"
    wsLog.Range(wsLog.Cells(r, baseCol), wsLog.Cells(r, baseCol + 2)).Font.Bold = True
    r = r + 1

    For i = 1 To issues.Count
        issueEntry = issues(i)
        wsLog.Cells(r, baseCol).Value = issueEntry(0)
        wsLog.Cells(r, baseCol + 1).Value = issueEntry(1)
        wsLog.Cells(r, baseCol + 2).Value = issueEntry(2)
        r = r + 1
    Next i

    WriteFCVIssueSection = r + 1
End Function

' ---------------------------------------------------------------------------
' Workbook name helpers
' ---------------------------------------------------------------------------

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
