Attribute VB_Name = "mod_09_ValidateSubfrConstraints"
Option Explicit

' Bookmark: Validate_Subframe_Constraints_v10
' Status: FIXED FOR host + LOG OUTPUT
' Target: Excel 2024 LTSC

Public Sub Validate_Subframe_Constraints()
    Dim startTime As Double
    startTime = Timer
    
    Dim wsExp As Worksheet
    Dim wsLog As Worksheet
    Dim wsPdu As Worksheet
    Dim tbl As ListObject
    Dim pduTbl As ListObject
    
    Dim data As Variant
    Dim aduData As Variant
    
    Dim i As Long, j As Long, m As Long, st As Long
    Dim startRow As Long, endRow As Long
    Dim currentSFN As Variant
    Dim maxSch As Long, nRx As Long
    
    Dim idxSFN As Long, idxTXID As Long, idxLEN As Long
    Dim rxColIdx() As Long
    
    Dim dictADU As Object
    Dim vCount As Long
    Dim writeRow As Long
    
    On Error GoTo CleanFail
    
    Set wsExp = ThisWorkbook.Worksheets("ExpResults")
    Set tbl = wsExp.ListObjects("ExpResultsTable")
    
    Set wsLog = Nothing
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("TX_SFN est Log")
    On Error GoTo CleanFail
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsLog.Name = "TX_SFN est Log"
    End If
    
    Set wsPdu = ThisWorkbook.Worksheets("PDU Size Table")
    Set pduTbl = wsPdu.ListObjects("ADU2NumSubchansTable")
    
    maxSch = GetWorkbookNameLong("Nsch_per_subfr")
    nRx = GetWorkbookNameLong("Num_Rx_Stations")
    
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "ExpResultsTable has no data rows.", vbInformation, "Validate Subframe Constraints"
        Exit Sub
    End If
    
    If pduTbl.DataBodyRange Is Nothing Then
        MsgBox "ADU2NumSubchansTable has no data rows.", vbExclamation, "Validate Subframe Constraints"
        Exit Sub
    End If
    
    data = tbl.DataBodyRange.Value
    aduData = pduTbl.DataBodyRange.Value
    
    Set dictADU = CreateObject("Scripting.Dictionary")
    dictADU.CompareMode = vbTextCompare
    
    For i = 1 To UBound(aduData, 1)
        If Not IsEmpty(aduData(i, 1)) Then
            dictADU(CStr(aduData(i, 1))) = aduData(i, 2)
        End If
    Next i
    
    idxSFN = tbl.ListColumns("TX_SFN_est").Index
    idxTXID = tbl.ListColumns("TX_ID").Index
    idxLEN = tbl.ListColumns("LEN").Index
    
    ReDim rxColIdx(1 To nRx)
    For i = 1 To nRx
        rxColIdx(i) = tbl.ListColumns("RXTIME" & CStr(i)).Index
    Next i
    
    wsLog.Range("I2:K100000").ClearContents
    
    With wsLog.Range("I2")
        .Value = "CONSTRAINT VALIDATION"
        .Font.Bold = True
        .Offset(1, 0).Value = "Timestamp:"
        .Offset(1, 1).Value = Now
        .Offset(2, 0).Value = "Issues Found:"
        .Offset(2, 1).Value = 0
        .Offset(3, 0).Value = "Processing (s):"
        .Offset(3, 1).Value = ""
        .Offset(5, 0).Value = "Type"
        .Offset(5, 1).Value = "Location"
        .Offset(5, 2).Value = "Description"
        .Offset(5, 0).Resize(1, 3).Font.Bold = True
    End With
    
    writeRow = 8
    vCount = 0
    
    Application.StatusBar = "Validate_Subframe_Constraints: scanning ExpResultsTable..."
    
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
        
        For st = 1 To nRx
            Dim rxSubSum As Long
            Dim isTX As Boolean
            Dim isRX As Boolean
            Dim rxCountForTx As Long
            
            rxSubSum = 0
            isTX = False
            isRX = False
            rxCountForTx = 0
            
            For j = startRow To endRow
                If IsNumeric(data(j, idxTXID)) Then
                    If CLng(data(j, idxTXID)) = st Then
                        isTX = True
                        For m = 1 To nRx
                            If Not IsEmpty(data(j, rxColIdx(m))) And Trim$(CStr(data(j, rxColIdx(m)))) <> "" Then
                                rxCountForTx = rxCountForTx + 1
                            End If
                        Next m
                    End If
                End If
                
                If Not IsEmpty(data(j, rxColIdx(st))) And Trim$(CStr(data(j, rxColIdx(st)))) <> "" Then
                    isRX = True
                    If dictADU.Exists(CStr(data(j, idxLEN))) Then
                        If IsNumeric(dictADU(CStr(data(j, idxLEN)))) Then
                            rxSubSum = rxSubSum + CLng(dictADU(CStr(data(j, idxLEN))))
                        End If
                    End If
                End If
            Next j
            
            If isTX And isRX Then
                vCount = vCount + 1
                wsLog.Cells(writeRow, "I").Value = IIf(rxCountForTx = 0, "GHOST CONFLICT", "HALF-DUPLEX")
                wsLog.Cells(writeRow, "J").Value = "SFN " & CStr(currentSFN)
                wsLog.Cells(writeRow, "K").Value = IIf(rxCountForTx = 0, _
                    "St " & st & " ghost TX during RX.", _
                    "St " & st & " TX/RX collision.")
                writeRow = writeRow + 1
            End If
            
            If rxSubSum > maxSch Then
                vCount = vCount + 1
                wsLog.Cells(writeRow, "I").Value = "CAPACITY"
                wsLog.Cells(writeRow, "J").Value = "SFN " & CStr(currentSFN)
                wsLog.Cells(writeRow, "K").Value = "St " & st & " sum " & rxSubSum & " > " & maxSch
                writeRow = writeRow + 1
            End If
        Next st
        
        i = i + 1
NextGroup:
    Loop
    
    wsLog.Range("J4").Value = vCount
    wsLog.Range("J5").Value = Round(Timer - startTime, 3)
    wsLog.Columns("I:K").AutoFit
    
    Application.StatusBar = "Validation Done: " & vCount & " issues in " & Round(Timer - startTime, 2) & " s"
    Exit Sub

CleanFail:
    Application.StatusBar = False
    MsgBox "Validate_Subframe_Constraints failed:" & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, _
           vbCritical, "Validate Subframe Constraints"
End Sub

Private Function GetWorkbookNameLong(ByVal nameText As String) As Long
    Dim nm As Name
    Dim expr As String
    Dim v As Variant
    
    On Error GoTo FailHard
    
    Set nm = ThisWorkbook.Names(nameText)
    expr = nm.RefersTo
    
    If Len(expr) > 0 Then
        If Left$(expr, 1) = "=" Then expr = Mid$(expr, 2)
    End If
    
    v = Application.Evaluate(expr)
    
    If IsError(v) Or Not IsNumeric(v) Then
        Err.Raise vbObjectError + 7100, "GetWorkbookNameLong", _
                  "Workbook name '" & nameText & "' did not evaluate to a numeric value."
    End If
    
    GetWorkbookNameLong = CLng(v)
    Exit Function

FailHard:
    Err.Raise vbObjectError + 7101, "GetWorkbookNameLong", _
              "Could not resolve workbook name: " & nameText
End Function

