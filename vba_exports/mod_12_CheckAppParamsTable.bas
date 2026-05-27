Attribute VB_Name = "mod_12_CheckAppParamsTable"
Sub CheckAppParamsTable()
    Dim wsParams As Worksheet, wsResults As Worksheet
    Dim srcSheet As Worksheet, srcWB As Workbook
    Dim targetTable As ListObject, appParamTable As ListObject
    Dim srcCols As Object, dictAppParams As Object
    Dim lastRow As Long, i As Long, missingCount As Long
    Dim srcData As Variant, appKey As String
    Dim appIDVal As Variant, lenVal As Variant, txIDVal As Variant
    Dim missingLog As String
    
    ' 1. INITIALIZE ENVIRONMENT & UTILITIES
    Set wsParams = ThisWorkbook.Sheets("Exp Config & Data Proc Params")
    Set wsResults = ThisWorkbook.Sheets("ExpResults")
    Set targetTable = wsResults.ListObjects("ExpResultsTable")
    
    On Error Resume Next
    Set appParamTable = wsParams.ListObjects("PDU2RXTprocVendorID")
    On Error GoTo 0
    
    If appParamTable Is Nothing Then
        MsgBox "Critical Error: 'PDU2RXTprocVendorID' ListObject table not found on parameters sheet!", vbCritical, "Validation Aborted"
        Exit Sub
    End If
    
    ' 2. INDEX VALID APP PARAMS (MAPPED PDU KEYS)
    ' Maps combined key: "LEN|VendorID" or "App_ID|VendorID" based on structural setup
    Set dictAppParams = CreateObject("Scripting.Dictionary")
    Dim paramData As Variant: paramData = appParamTable.Range.Value
    Dim r As Long, c As Long
    
    ' Loop through columns (Vendors) and rows (Lengths/Apps) to map known universal boundaries
    For c = 2 To UBound(paramData, 2)
        For r = 2 To UBound(paramData, 1)
            If Not IsEmpty(paramData(r, 1)) And Not IsEmpty(paramData(1, c)) Then
                ' Create cross-reference key profile lookup
                appKey = UCase(Trim(paramData(r, 1))) & "|" & UCase(Trim(paramData(1, c)))
                dictAppParams(appKey) = True
            End If
        Next r
    Next c

    ' 3. MAP STATION ID TO VENDOR ID (To identify which vendor profile an App relies on)
    Dim dictS2V As Object: Set dictS2V = CreateObject("Scripting.Dictionary")
    Dim vArr As Variant: vArr = wsParams.Range("StationID2VendorID").Value
    For i = 1 To UBound(vArr, 1)
        If Not IsEmpty(vArr(i, 1)) Then dictS2V(UCase(Trim(CStr(vArr(i, 1))))) = UCase(Trim(vArr(i, 2)))
    Next i

    ' 4. RETRIEVE ACTIVE RAW LOG DATA SINK
    ' Accessing active data array to check what data fields are dropping in
    On Error Resume Next
    Set srcWB = ActiveWorkbook
    Set srcSheet = srcWB.Sheets(1)
    On Error GoTo 0
    
    If srcWB.Name = ThisWorkbook.Name Then
        MsgBox "Please select or focus the external raw data file before running parameter validation.", vbExclamation, "No Data Source Found"
        Exit Sub
    End If

    lastRow = srcSheet.Cells(srcSheet.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "The selected data source sheet appears to contain no records.", vbExclamation, "Empty Dataset"
        Exit Sub
    End If

    ' Map dynamic runtime raw headings
    Set srcCols = CreateObject("Scripting.Dictionary")
    For i = 1 To srcSheet.Cells(1, srcSheet.Columns.count).End(xlToLeft).Column
        srcCols(UCase(Trim(srcSheet.Cells(1, i).Value))) = i
    Next i

    ' Validate necessary column headers exist in the target raw logs
    If Not (srcCols.Exists("LEN") And srcCols.Exists("TX_ID") And (srcCols.Exists("IVI_ID") Or srcCols.Exists("APP_ID"))) Then
        MsgBox "Missing critical mapping column headers inside the raw file." & vbCrLf & _
               "Ensure 'LEN', 'TX_ID', and 'IVI_ID' (or 'APP_ID') are present.", vbCritical, "Header Mapping Failed"
        Exit Sub
    End If

    ' Extract raw array into memory for rapid loop scanning
    srcData = srcSheet.Range(srcSheet.Cells(2, 1), srcSheet.Cells(lastRow, srcSheet.Cells(1, srcSheet.Columns.count).End(xlToLeft).Column)).Value
    
    ' 5. RUN PRE-FLIGHT AUDIT LOOP
    Dim idxAppCol As Long
    If srcCols.Exists("IVI_ID") Then idxAppCol = srcCols("IVI_ID") Else idxAppCol = srcCols("APP_ID")
    
    Dim idxLenCol As Long: idxLenCol = srcCols("LEN")
    Dim idxTxCol As Long: idxTxCol = srcCols("TX_ID")
    
    Dim missingDict As Object: Set missingDict = CreateObject("Scripting.Dictionary")
    missingCount = 0
    
    For r = 1 To UBound(srcData, 1)
        appIDVal = UCase(Trim(CStr(srcData(r, idxAppCol))))
        lenVal = UCase(Trim(CStr(srcData(r, idxLenCol))))
        txIDVal = UCase(Trim(CStr(srcData(r, idxTxCol))))
        
        ' Identify the vendor using the station ID cross-reference table
        Dim currentVendor As String: currentVendor = ""
        If dictS2V.Exists(txIDVal) Then currentVendor = dictS2V(txIDVal)
        
        ' Build composite lookup key matching engine specifications: "LEN|VendorID"
        Dim diagnosticKey As String: diagnosticKey = lenVal & "|" & currentVendor
        
        If currentVendor <> "" Then
            ' If the combined specification parameter is missing, log it as an unmapped profile
            If Not dictAppParams.Exists(diagnosticKey) Then
                Dim errorReportString As String
                errorReportString = "App/IVI ID: " & appIDVal & " (Len: " & lenVal & " bytes for Vendor: " & currentVendor & ")"
                
                If Not missingDict.Exists(errorReportString) Then
                    missingDict(errorReportString) = 1
                    missingCount = missingCount + 1
                Else
                    missingDict(errorReportString) = missingDict(errorReportString) + 1
                End If
            End If
        Else
            ' Handle completely undocumented station transmissions
            Dim missingStationString As String
            missingStationString = "STATION ID NOT MAPPED: " & txIDVal & " (App ID: " & appIDVal & ", Len: " & lenVal & ")"
            If Not missingDict.Exists(missingStationString) Then
                missingDict(missingStationString) = 1
                missingCount = missingCount + 1
            Else
                missingDict(missingStationString) = missingDict(missingStationString) + 1
            End If
        End If
    Next r

    ' 6. DISSEMINATE COMPLIANCE AUDIT METRICS
    If missingCount = 0 Then
        MsgBox "Configuration Audit passed successfully!" & vbCrLf & vbCrLf & _
               "All " & Format(UBound(srcData, 1), "#,##0") & " rows match valid parameter bounds in the 'PDU2RXTprocVendorID' table.", _
               vbInformation, "Parameter Mapping Compliance: 100%"
    Else
        missingLog = "The following configuration mappings are missing from your App Params tables:" & vbCrLf & vbCrLf
        
        Dim alertKey As Variant, trackingIdx As Long: trackingIdx = 0
        For Each alertKey In missingDict.Keys
            trackingIdx = trackingIdx + 1
            missingLog = missingLog & " • " & alertKey & " ? Found in " & missingDict(alertKey) & " log entries." & vbCrLf
            ' Prevent infinite window scroll heights inside VBA message limits
            If trackingIdx >= 20 Then
                missingLog = missingLog & " • [...and " & (missingDict.count - 20) & " more unmapped structural definitions]" & vbCrLf
                Exit For
            End If
        Next alertKey
        
        MsgBox missingLog, vbCritical, "Audit Failure: " & missingDict.count & " Unmapped Configurations Detected"
    End If
End Sub
