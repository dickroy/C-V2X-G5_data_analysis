Option Explicit

' Version: V1.0.2-DEBUG
' DEBUG build of TX_SFN conflict resolution.
' Rule: GROUPS SHALL NEVER BE SPLIT.
' Purpose: trace every pool/cset decision and row move to diagnose why valid groups were altered.

Private Const MODULE_VERSION_TXSFNCR As String = "V1.0.2-DEBUG"
Private Const DEBUG_TXSFNCR As Boolean = True
Private Const DEBUG_VERBOSE_TXSFNCR As Boolean = True
Private Const DEBUG_LOG_SHEET As String = "Conflict Resolution Debug Log"
Private Const DEBUG_WARN_ROW_THRESHOLD As Long = 100
Private Const NO_RX_TIME As Double = -1#

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter_TXSFNCR Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency_TXSFNCR Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter_TXSFNCR Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency_TXSFNCR Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
#End If

Private Type TCsetAnalysis
    MovesRequired As Long
    candidateCount As Long
    candidateRows() As Long
End Type

Private mData As Variant
Private mFilteredCount As Long
Private mIdxTXID As Long
Private mIdxTXQ As Long
Private mIdxSFNCol As Long
Private mIdxLEN As Long
Private mIdxTXperSFN As Long
Private mIdxGen As Long
Private mIdxAvg As Long
Private mIdxTotLat As Long
Private mIdxRxCnt As Long
Private mTxBitmap As String
Private mBitmapLen As Long
Private mRxStationIDs() As Long
Private mRxDataColIdx() As Long
Private mActiveRxCount As Long
Private mDictS2V As Object
Private mDictVC As Object
Private mDictA2P As Object
Private mDictP2R As Object
Private mDictP2Sigma As Object
Private mInitialSFN() As Long
Private mCurrentSFN() As Long
Private mRowTXID() As Long
Private mRowTXQTime() As Double
Private mRowNsch() As Long
Private mRowPduKey() As String
Private mRowMinRxTime() As Double
Private mRowOriginalIndex() As Long
Private mRowValidInput() As Boolean
Private mRowRXCount() As Long
Private mWritten() As Boolean
Private mPoolRows() As Long
Private mPoolCount As Long
Private mPoolMinSFN As Long
Private mPoolMaxSFN As Long
Private mPoolCenter As Double
Private mPoolCestStartRows() As Long
Private mPoolCestEndRows() As Long
Private mPoolCestSFN() As Long
Private mPoolCestCount As Long
Private mOutputData() As Variant
Private mOutputCount As Long
Private mOutputWritePos As Long
Private mScanPos As Long
Private mPoolCountResolved As Long
Private mMaxObservedPoolSize As Long
Private mRemainingViolations As Long
Private mUnresolvedAttemptCount As Long
Private mDiagCount As Long
Private mFindConflictSeconds As Double
Private mBuildPoolSeconds As Double
Private mResolvePoolSeconds As Double
Private mWritePoolSeconds As Double
Private mFinalizeSeconds As Double
Private mDebugWs As Worksheet
Private mDebugRow As Long
Private mOriginalConflictCount As Long

Public Sub Run_TX_SFNConflictResolution_DEBUG()
    TX_SFNConflictResolution_DEBUG
End Sub

Public Sub TX_SFNConflictResolution_DEBUG( _
    Optional ByRef data As Variant, Optional ByVal filteredCount As Long = 0, Optional ByVal idxSFNCol As Long = 0, Optional ByVal idxTXID As Long = 0, _
    Optional ByVal idxTXQ As Long = 0, Optional ByVal idxLEN As Long = 0, Optional ByVal idxTXperSFN As Long = 0, Optional ByVal idxRxCnt As Long = 0, _
    Optional ByVal idxAvg As Long = 0, Optional ByVal idxTotLat As Long = 0, Optional ByVal idxGen As Long = 0, Optional ByRef rxDataColIdx() As Long, _
    Optional ByRef rxStationIDs() As Long, Optional ByVal activeRxCount As Long = 0, Optional ByRef dictS2V As Object, Optional ByRef dictVC As Object, _
    Optional ByRef dictA2P As Object, Optional ByRef dictP2R As Object, Optional ByRef dictP2Sigma As Object, Optional ByVal txBitmap As String = "", _
    Optional ByVal bitmapLen As Long = 0, Optional ByRef elapsedSeconds As Double)

    Dim startTime As Double
    Dim conflictStart As Long
    Dim poolMoved As Boolean
    Dim continueAllowed As VbMsgBoxResult

    If filteredCount > DEBUG_WARN_ROW_THRESHOLD Then
        continueAllowed = MsgBox("WARNING: This routine creates copious outputs." & vbCrLf & vbCrLf & _
                                 "Are you sure you want to continue?" & vbCrLf & vbCrLf & _
                                 "Rows to be processed: " & filteredCount, _
                                 vbYesNo + vbExclamation, _
                                 "TX_SFN Conflict Resolution DEBUG")
        If continueAllowed = vbNo Then Exit Sub
    End If

    startTime = MicroTimer_TXSFNCR()
    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, txBitmap, bitmapLen
    SetupDebugLog
    LogDebugHeader

    If Not ValidateInputMonotoneTXSFN() Then
        LogDebug "ABORT", "Input SFN sequence is not monotone.", "", "", "", ""
        GoTo CleanExit
    End If

    PrepareRowDerivedData
    InitializeOutputBuffer
    LogDebug "INIT", "rows=" & mFilteredCount, "bitmapLen=" & mBitmapLen, "activeRx=" & mActiveRxCount, "", ""

    Do
        conflictStart = FindNextConflictStart()
        If conflictStart <= 0 Then Exit Do

        BuildPoolFromConflictStart conflictStart
        LogPoolState conflictStart

        poolMoved = ResolveEntirePool()
        LogDebug "POOL_RESOLVE_COMPLETE", "moved=" & CStr(poolMoved), "poolCount=" & mPoolCount, "resolvedCount=" & mPoolCountResolved, "unresolvedAttempts=" & mUnresolvedAttemptCount, ""

        WriteResolvedPoolToOutput
        If mPoolCount > 0 Then mScanPos = mPoolRows(mPoolCount) + 1
    Loop

    WriteUnwrittenRowsToOutput
    RecomputeFinalTXperSFN
    ValidateResolvedRXTimingOnly
    FinalizeOutputVariant data
    elapsedSeconds = MicroTimer_TXSFNCR() - startTime
    LogDebug "DONE", "elapsedSeconds=" & Format$(elapsedSeconds, "0.000"), "remainingViolations=" & mRemainingViolations, "resolved=" & mPoolCountResolved, "unresolvedAttempts=" & mUnresolvedAttemptCount, ""

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Private Sub SetupDebugLog()
    On Error Resume Next
    Set mDebugWs = ThisWorkbook.Worksheets(DEBUG_LOG_SHEET)
    On Error GoTo 0
    If mDebugWs Is Nothing Then
        Set mDebugWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        mDebugWs.Name = DEBUG_LOG_SHEET
    End If
    mDebugWs.Cells.Clear
    mDebugRow = 1
End Sub

Private Sub LogDebugHeader()
    mDebugWs.Cells(mDebugRow, 1).Value = "TX_SFN Conflict Resolution DEBUG"
    mDebugWs.Cells(mDebugRow, 2).Value = MODULE_VERSION_TXSFNCR
    mDebugRow = mDebugRow + 2
    mDebugWs.Cells(mDebugRow, 1).Value = "RULE"
    mDebugWs.Cells(mDebugRow, 2).Value = "GROUPS SHALL NEVER BE SPLIT"
    mDebugRow = mDebugRow + 2
    mDebugWs.Cells(mDebugRow, 1).Resize(1, 6).Value = Array("Event", "A", "B", "C", "D", "E")
    mDebugRow = mDebugRow + 1
End Sub

Private Sub LogDebug(ByVal eventName As String, ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As String)
    If mDebugWs Is Nothing Then Exit Sub
    mDebugWs.Cells(mDebugRow, 1).Value = eventName
    mDebugWs.Cells(mDebugRow, 2).Value = a
    mDebugWs.Cells(mDebugRow, 3).Value = b
    mDebugWs.Cells(mDebugRow, 4).Value = c
    mDebugWs.Cells(mDebugRow, 5).Value = d
    mDebugWs.Cells(mDebugRow, 6).Value = e
    mDebugRow = mDebugRow + 1
End Sub

Private Sub LogPoolState(ByVal conflictStart As Long)
    Dim i As Long, s As String
    s = "rows="
    For i = 1 To mPoolCount
        s = s & mPoolRows(i) & IIf(i < mPoolCount, ",", "")
    Next i
    LogDebug "POOL_BUILT", "seedConflictStart=" & conflictStart, "poolCount=" & mPoolCount, "minSFN=" & mPoolMinSFN, "maxSFN=" & mPoolMaxSFN, s
End Sub

Private Function MicroTimer_TXSFNCR() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency
    If QueryPerformanceFrequency_TXSFNCR(cyFreq) <> 0 Then
        QueryPerformanceCounter_TXSFNCR cyTicks
        If cyFreq > 0 Then MicroTimer_TXSFNCR = cyTicks / cyFreq
    End If
End Function

Private Sub InitializeContext(ByRef data As Variant, ByVal filteredCount As Long, ByVal idxSFNCol As Long, ByVal idxTXID As Long, ByVal idxTXQ As Long, ByVal idxLEN As Long, ByVal idxTXperSFN As Long, ByVal idxRxCnt As Long, ByVal idxAvg As Long, ByVal idxTotLat As Long, ByVal idxGen As Long, ByRef rxDataColIdx() As Long, ByRef rxStationIDs() As Long, ByVal activeRxCount As Long, ByRef dictS2V As Object, ByRef dictVC As Object, ByRef dictA2P As Object, ByRef dictP2R As Object, ByRef dictP2Sigma As Object, ByVal txBitmap As String, ByVal bitmapLen As Long)
    mData = data
    mFilteredCount = filteredCount
    mIdxSFNCol = idxSFNCol
    mIdxTXID = idxTXID
    mIdxTXQ = idxTXQ
    mIdxLEN = idxLEN
    mIdxTXperSFN = idxTXperSFN
    mIdxRxCnt = idxRxCnt
    mIdxAvg = idxAvg
    mIdxTotLat = idxTotLat
    mIdxGen = idxGen
    mRxDataColIdx = rxDataColIdx
    mRxStationIDs = rxStationIDs
    mActiveRxCount = activeRxCount
    Set mDictS2V = dictS2V
    Set mDictVC = dictVC
    Set mDictA2P = dictA2P
    Set mDictP2R = dictP2R
    Set mDictP2Sigma = dictP2Sigma
    mTxBitmap = txBitmap
    mBitmapLen = bitmapLen
    mScanPos = 1
    mOutputWritePos = 1
    mPoolCount = 0
    mPoolCestCount = 0
    mOriginalConflictCount = 0
End Sub

' [rest of module unchanged]
