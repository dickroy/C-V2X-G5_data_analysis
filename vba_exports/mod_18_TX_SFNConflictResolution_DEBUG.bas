Attribute VB_Name = "mod_18_TX_SFNConflictResolution_DEBUG"
Option Explicit

' Version: V1.0.4-DEBUG
' Debug module based on mod_17_TX_SFNConflictResolution behavior.
Private Const MODULE_VERSION_TXSFNCR_DEBUG As String = "V1.0.4-DEBUG"
Private Const DEBUG_TXSFNCR As Boolean = True
Private Const PROGRESS_STEP_ROWS As Long = 10000
Private Const DEBUG_PROMPT_ROW_LIMIT As Long = 100

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

Private Const NO_RX_TIME As Double = -1#

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
Private mMissingPduSizes As Object
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
Private mResolveEntirePoolSeconds As Double
Private mResolveOneCsetSeconds As Double
Private mTryPlaceMoveSeconds As Double
Private mLegalCheckSeconds As Double
Private mBucketExcludeSeconds As Double
Private mBucketAddSeconds As Double
Private mValidateBucketSeconds As Double
Private mResolveOneCsetExtractSeconds As Double
Private mResolveOneCsetPostSeconds As Double
Private mDebugSheet As Worksheet
Private mDebugRow As Long
Private mDebugSeq As Long
Private mDebugStartTime As Double
Private mCurrentCsetIdx As Long
Private mLastAttemptReason As String

Public Sub TX_SFNConflictResolution_DEBUG()
    Dim wsExp As Worksheet
    Dim lo As ListObject
    Dim data As Variant
    Dim filteredCount As Long
    Dim elapsedSeconds As Double
    Dim idxSFNCol As Long
    Dim idxTXID As Long
    Dim idxTXQ As Long
    Dim idxLEN As Long
    Dim idxTXperSFN As Long
    Dim idxRxCnt As Long
    Dim idxAvg As Long
    Dim idxTotLat As Long
    Dim idxGen As Long
    Dim rxDataColIdx() As Long
    Dim rxStationIDs() As Long
    Dim activeRxCount As Long
    Dim dictS2V As Object
    Dim dictVC As Object
    Dim dictA2P As Object
    Dim dictP2R As Object
    Dim dictP2Sigma As Object
    Dim txBitmap As String
    Dim bitmapLen As Long
    Dim promptResult As VbMsgBoxResult

    On Error GoTo HandleError

    Set wsExp = ThisWorkbook.Worksheets("ExpResults")
    Set lo = wsExp.ListObjects("ExpResultsTable")
    If lo.DataBodyRange Is Nothing Then
        MsgBox "ExpResultsTable has no rows to process.", vbInformation, "TX_SFN Conflict Resolution DEBUG " & MODULE_VERSION_TXSFNCR_DEBUG
        Exit Sub
    End If

    filteredCount = lo.DataBodyRange.Rows.Count
    If filteredCount > DEBUG_PROMPT_ROW_LIMIT Then
        promptResult = MsgBox("WARNING: This routine creates copious outputs. Are you sure you want to continue?", vbYesNo + vbExclamation, "TX_SFN Conflict Resolution DEBUG " & MODULE_VERSION_TXSFNCR_DEBUG)
        If promptResult <> vbYes Then Exit Sub
    End If

    idxTXID = GetRequiredColumnIndex(lo, "TX_ID")
    idxTXQ = GetRequiredColumnIndex(lo, "TXQTIME")
    idxSFNCol = GetRequiredColumnIndex(lo, "TX_SFN_est")
    idxLEN = GetRequiredColumnIndex(lo, "LEN")
    idxTXperSFN = GetRequiredColumnIndex(lo, "TXperSFN")
    idxGen = GetRequiredColumnIndex(lo, "MSG_GEN_TIME")
    idxAvg = GetRequiredColumnIndex(lo, "AVG TX1 RXTIMES")
    idxTotLat = GetRequiredColumnIndex(lo, "AVG_TOTAL_LATENCY")
    idxRxCnt = GetRequiredColumnIndex(lo, "RX_COUNT")

    BuildRxColumnMaps lo, rxDataColIdx, rxStationIDs, activeRxCount
    txBitmap = GetConfigText("tx_bitmap")
    bitmapLen = GetConfigLong("bitmap_len")

    Set dictS2V = CreateObject("Scripting.Dictionary")
    Set dictVC = CreateObject("Scripting.Dictionary")
    Set dictA2P = CreateObject("Scripting.Dictionary")
    Set dictP2R = CreateObject("Scripting.Dictionary")
    Set dictP2Sigma = CreateObject("Scripting.Dictionary")

    data = lo.DataBodyRange.Value2

    TX_SFNConflictResolution_DEBUG_Core data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, _
                                        idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, _
                                        activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, _
                                        txBitmap, bitmapLen, elapsedSeconds

    lo.DataBodyRange.Value2 = data
    MsgBox "TX_SFNConflictResolution_DEBUG completed in " & Format$(elapsedSeconds, "0.000") & " s.", vbInformation, "TX_SFN Conflict Resolution DEBUG " & MODULE_VERSION_TXSFNCR_DEBUG
    Exit Sub

HandleError:
    MsgBox "TX_SFNConflictResolution_DEBUG failed: " & Err.Description, vbCritical, "TX_SFN Conflict Resolution DEBUG " & MODULE_VERSION_TXSFNCR_DEBUG
End Sub

Private Sub TX_SFNConflictResolution_DEBUG_Core( _
    ByRef data As Variant, ByVal filteredCount As Long, ByVal idxSFNCol As Long, ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, ByVal idxLEN As Long, ByVal idxTXperSFN As Long, ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, ByVal idxTotLat As Long, ByVal idxGen As Long, ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, ByVal activeRxCount As Long, ByRef dictS2V As Object, ByRef dictVC As Object, _
    ByRef dictA2P As Object, ByRef dictP2R As Object, ByRef dictP2Sigma As Object, ByVal txBitmap As String, _
    ByVal bitmapLen As Long, ByRef elapsedSeconds As Double)

    Dim startTime As Double
    Dim t0 As Double
    Dim conflictStart As Long
    Dim totalFindSeconds As Double
    Dim totalBuildSeconds As Double
    Dim totalResolveSeconds As Double
    Dim totalWriteSeconds As Double
    Dim totalFinalSeconds As Double
    Dim totalLoopCount As Long
    Dim oldStatusBar As Variant
    Dim tPhase As Double

    startTime = MicroTimer_TXSFNCR()
    mDebugStartTime = startTime
    EnsureDebugLogSheet
    DebugLog "INIT", "Version", MODULE_VERSION_TXSFNCR_DEBUG, "Rows", filteredCount, "BitmapLen", bitmapLen, "ActiveRx", activeRxCount
    oldStatusBar = Application.StatusBar
    Application.StatusBar = "TX_SFN conflict resolution running..."

    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, txBitmap, bitmapLen

    If Not ValidateInputMonotoneTXSFN() Then GoTo CleanExit

    PrepareRowDerivedData
    InitializeOutputBuffer
    mTxBitmap = txBitmap

    If Not IsAllOnesBitmap(mTxBitmap) Then
        MsgBox "TX_SFN Conflict Resolution skipped: TX bitmap is not all ones.", vbInformation, "TX_SFN Conflict Resolution"
        GoTo CleanExit
    End If

    Do
        t0 = MicroTimer_TXSFNCR()

        tPhase = MicroTimer_TXSFNCR()
        conflictStart = FindNextConflictStart()
        totalFindSeconds = totalFindSeconds + (MicroTimer_TXSFNCR() - tPhase)

        If conflictStart <= 0 Then Exit Do
        DebugLog "CONFLICT_START", "Row", conflictStart, "SFN", mCurrentSFN(conflictStart), "NextSFN", mCurrentSFN(conflictStart + 1)

        tPhase = MicroTimer_TXSFNCR()
        BuildPoolFromConflictStart conflictStart
        totalBuildSeconds = totalBuildSeconds + (MicroTimer_TXSFNCR() - tPhase)

        tPhase = MicroTimer_TXSFNCR()
        ResolveEntirePool
        totalResolveSeconds = totalResolveSeconds + (MicroTimer_TXSFNCR() - tPhase)

        tPhase = MicroTimer_TXSFNCR()
        WriteResolvedPoolToOutput
        totalWriteSeconds = totalWriteSeconds + (MicroTimer_TXSFNCR() - tPhase)

        If mPoolCount > 0 Then
            mScanPos = mPoolRows(mPoolCount) + 1
        End If

        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR() - t0)
        totalLoopCount = totalLoopCount + 1
    Loop

    t0 = MicroTimer_TXSFNCR()
    WriteUnwrittenRowsToOutput
    RecomputeFinalTXperSFN
    ValidateResolvedRXTimingOnly
    FinalizeOutputVariant data
    totalFinalSeconds = MicroTimer_TXSFNCR() - t0

    mFinalizeSeconds = mFinalizeSeconds + totalFinalSeconds
    elapsedSeconds = MicroTimer_TXSFNCR() - startTime

    Debug.Print "SUMMARY", "loops", totalLoopCount, "find", Format$(totalFindSeconds, "0.000"), "build", Format$(totalBuildSeconds, "0.000"), "resolve", Format$(totalResolveSeconds, "0.000"), "write", Format$(totalWriteSeconds, "0.000"), "final", Format$(totalFinalSeconds, "0.000"), "total", Format$(elapsedSeconds, "0.000")
    DebugLog "SUMMARY", "ElapsedSec", Format$(elapsedSeconds, "0.000"), "ResolvedPools", mPoolCountResolved, "UnresolvedAttempts", mUnresolvedAttemptCount, "RemainingViolations", mRemainingViolations, "Loops", totalLoopCount

    WriteTimingResultsToConflictResolutionLog totalLoopCount, totalFindSeconds, totalBuildSeconds, totalResolveSeconds, totalWriteSeconds, totalFinalSeconds, elapsedSeconds
    WriteDebugSummary totalLoopCount, totalFindSeconds, totalBuildSeconds, totalResolveSeconds, totalWriteSeconds, totalFinalSeconds, elapsedSeconds

CleanExit:
    Application.StatusBar = oldStatusBar
End Sub

Private Sub WriteTimingResultsToConflictResolutionLog( _
    ByVal totalLoopCount As Long, _
    ByVal totalFindSeconds As Double, _
    ByVal totalBuildSeconds As Double, _
    ByVal totalResolveSeconds As Double, _
    ByVal totalWriteSeconds As Double, _
    ByVal totalFinalSeconds As Double, _
    ByVal elapsedSeconds As Double)

    Dim ws As Worksheet
    Dim r As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Conflict Resolution Log")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Range("A2:B50").ClearContents

    r = 2

    ws.Cells(r, 1).Value = "TX_SFN CR Timing"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    ws.Cells(r, 1).Value = "loops": ws.Cells(r, 2).Value = totalLoopCount: r = r + 1
    ws.Cells(r, 1).Value = "find": ws.Cells(r, 2).Value = totalFindSeconds: r = r + 1
    ws.Cells(r, 1).Value = "build": ws.Cells(r, 2).Value = totalBuildSeconds: r = r + 1
    ws.Cells(r, 1).Value = "resolve": ws.Cells(r, 2).Value = totalResolveSeconds: r = r + 1
    ws.Cells(r, 1).Value = "write": ws.Cells(r, 2).Value = totalWriteSeconds: r = r + 1
    ws.Cells(r, 1).Value = "final": ws.Cells(r, 2).Value = totalFinalSeconds: r = r + 1
    ws.Cells(r, 1).Value = "total": ws.Cells(r, 2).Value = elapsedSeconds: r = r + 1
    ws.Cells(r, 1).Value = "ResolveEntirePool": ws.Cells(r, 2).Value = mResolveEntirePoolSeconds: r = r + 1
    ws.Cells(r, 1).Value = "ResolveOneCset": ws.Cells(r, 2).Value = mResolveOneCsetSeconds: r = r + 1
    ws.Cells(r, 1).Value = "TryPlaceMove": ws.Cells(r, 2).Value = mTryPlaceMoveSeconds: r = r + 1
    ws.Cells(r, 1).Value = "LegalCheck": ws.Cells(r, 2).Value = mLegalCheckSeconds: r = r + 1
    ws.Cells(r, 1).Value = "BucketExclude": ws.Cells(r, 2).Value = mBucketExcludeSeconds: r = r + 1
    ws.Cells(r, 1).Value = "BucketAdd": ws.Cells(r, 2).Value = mBucketAddSeconds: r = r + 1
    ws.Cells(r, 1).Value = "ValidateBucket": ws.Cells(r, 2).Value = mValidateBucketSeconds: r = r + 1
    ws.Cells(r, 1).Value = "ResolveOneCsetExtract": ws.Cells(r, 2).Value = mResolveOneCsetExtractSeconds: r = r + 1
    ws.Cells(r, 1).Value = "ResolveOneCsetPost": ws.Cells(r, 2).Value = mResolveOneCsetPostSeconds
End Sub

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
    mCurrentCsetIdx = 0
    mLastAttemptReason = vbNullString
    If DEBUG_TXSFNCR Then Debug.Print "TX_SFNCR init: filteredCount=" & mFilteredCount & " bitmapLen=" & mBitmapLen
End Sub

Private Function ValidateInputMonotoneTXSFN() As Boolean
    Dim r As Long
    Dim prevVal As Long
    Dim curVal As Long

    ValidateInputMonotoneTXSFN = True
    If mFilteredCount <= 1 Then Exit Function

    prevVal = CLng(mData(1, mIdxSFNCol))
    For r = 2 To mFilteredCount
        curVal = CLng(mData(r, mIdxSFNCol))
        If curVal < prevVal Then
            ValidateInputMonotoneTXSFN = False
            Exit Function
        End If
        prevVal = curVal
    Next r
End Function

Private Sub PrepareRowDerivedData()
    Dim r As Long

    ReDim mInitialSFN(1 To mFilteredCount)
    ReDim mCurrentSFN(1 To mFilteredCount)
    ReDim mRowTXID(1 To mFilteredCount)
    ReDim mRowTXQTime(1 To mFilteredCount)
    ReDim mRowNsch(1 To mFilteredCount)
    ReDim mRowPduKey(1 To mFilteredCount)
    ReDim mRowMinRxTime(1 To mFilteredCount)
    ReDim mRowOriginalIndex(1 To mFilteredCount)
    ReDim mRowValidInput(1 To mFilteredCount)
    ReDim mRowRXCount(1 To mFilteredCount)
    ReDim mWritten(1 To mFilteredCount)

    For r = 1 To mFilteredCount
        mInitialSFN(r) = CLng(mData(r, mIdxSFNCol))
        mCurrentSFN(r) = mInitialSFN(r)
        mRowTXID(r) = CLng(mData(r, mIdxTXID))
        mRowTXQTime(r) = CDbl(mData(r, mIdxTXQ))
        mRowNsch(r) = CLng(mData(r, mIdxLEN))
        mRowPduKey(r) = CStr(mData(r, mIdxLEN))
        mRowMinRxTime(r) = GetRowMinRxTime(r)
        mRowOriginalIndex(r) = r
        mRowValidInput(r) = True
        If mIdxRxCnt > 0 Then
            mRowRXCount(r) = CLng(mData(r, mIdxRxCnt))
        Else
            mRowRXCount(r) = 0
        End If
    Next r
End Sub

Private Function FindNextConflictStart() As Long
    Dim r As Long
    For r = mScanPos To mFilteredCount - 1
        If mCurrentSFN(r) = mCurrentSFN(r + 1) Then
            FindNextConflictStart = r
            DebugLog "CONFLICT_FOUND", "Row", r, "SFN", mCurrentSFN(r), "NextRow", r + 1
            Exit Function
        End If
        If (r Mod PROGRESS_STEP_ROWS) = 0 Then UpdateProgressBar r
    Next r
End Function

Private Sub BuildPoolFromConflictStart(ByVal startRow As Long)
    Dim leftRow As Long
    Dim rightRow As Long
    Dim i As Long

    If startRow < 1 Or startRow >= mFilteredCount Then Exit Sub

    leftRow = startRow
    rightRow = startRow + 1

    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = mCurrentSFN(leftRow)
        leftRow = leftRow - 1
    Loop

    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = mCurrentSFN(rightRow)
        rightRow = rightRow + 1
    Loop

    mPoolCount = rightRow - leftRow + 1
    If mPoolCount <= 0 Then Exit Sub

    ReDim mPoolRows(1 To mPoolCount)
    For i = 1 To mPoolCount
        mPoolRows(i) = leftRow + i - 1
    Next i

    mPoolMinSFN = mCurrentSFN(leftRow)
    mPoolMaxSFN = mCurrentSFN(rightRow)
    mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#

    If mPoolCount > mMaxObservedPoolSize Then
        mMaxObservedPoolSize = mPoolCount
    End If

    If DEBUG_TXSFNCR Then
        Debug.Print "POOL built", "startRow", startRow, "leftRow", leftRow, "rightRow", rightRow, "poolCount", mPoolCount, "minSFN", mPoolMinSFN, "maxSFN", mPoolMaxSFN
    End If
    DebugLog "POOL_BUILD", "SeedConflictRow", startRow, "LeftRow", leftRow, "RightRow", rightRow, "PoolCount", mPoolCount, "PoolRows", BuildPoolRowListText(), "SFNRange", mPoolMinSFN & "..." & mPoolMaxSFN

    BuildPoolCests
End Sub

Private Sub BuildPoolCests()
    Dim i As Long
    Dim r As Long
    Dim startIdx As Long
    Dim curSFN As Long

    If mPoolCount <= 0 Then Exit Sub

    ReDim mPoolCestStartRows(1 To mPoolCount)
    ReDim mPoolCestEndRows(1 To mPoolCount)
    ReDim mPoolCestSFN(1 To mPoolCount)

    mPoolCestCount = 0
    i = 1

    Do While i <= mPoolCount
        startIdx = i
        r = mPoolRows(i)
        curSFN = mCurrentSFN(r)

        Do While i < mPoolCount
            If mCurrentSFN(mPoolRows(i + 1)) <> curSFN Then Exit Do
            i = i + 1
        Loop

        mPoolCestCount = mPoolCestCount + 1
        mPoolCestStartRows(mPoolCestCount) = startIdx
        mPoolCestEndRows(mPoolCestCount) = i
        mPoolCestSFN(mPoolCestCount) = curSFN

        If DEBUG_TXSFNCR Then
            Debug.Print "CEST chunk", "idx", mPoolCestCount, "rows", startIdx & ".." & i, "sfn", curSFN, "count", (i - startIdx + 1)
        End If
        DebugLog "CSET_CHUNK", "CsetIdx", mPoolCestCount, "PoolIndexRange", startIdx & ".." & i, "SFN", curSFN, "RowCount", (i - startIdx + 1), "Rows", BuildRowListFromPoolRange(startIdx, i)

        i = i + 1
    Loop
End Sub

Private Function ResolveEntirePool() As Boolean
    Dim cestIdx As Long
    Dim orderedCests() As Long
    Dim i As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    ResolveEntirePool = False

    If mPoolCestCount <= 0 Then
        mResolveEntirePoolSeconds = mResolveEntirePoolSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    orderedCests = GetOrderedCestIndexes()
    DebugLog "POOL_ANALYZE", "CsetCount", mPoolCestCount, "OrderedCsets", JoinLongArray(orderedCests)

    For i = LBound(orderedCests) To UBound(orderedCests)
        cestIdx = orderedCests(i)
        If cestIdx > 0 Then
            If ResolveOneCset(cestIdx) Then
                ResolveEntirePool = True
            End If
        End If
    Next i

    mResolveEntirePoolSeconds = mResolveEntirePoolSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function ResolveOneCset(ByVal cestIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim changed As Boolean
    Dim movedRowIdx As Long
    Dim sourceSFN As Long
    Dim t As Double
    Dim tExtract As Double
    Dim tPost As Double

    t = MicroTimer_TXSFNCR()
    ResolveOneCset = False
    mCurrentCsetIdx = cestIdx

    If cestIdx < 1 Or cestIdx > mPoolCestCount Then
        mResolveOneCsetSeconds = mResolveOneCsetSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    rowCount = mPoolCestEndRows(cestIdx) - mPoolCestStartRows(cestIdx) + 1
    If rowCount <= 1 Then
        mResolveOneCsetSeconds = mResolveOneCsetSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    tExtract = MicroTimer_TXSFNCR()
    rows = ExtractPoolRows(mPoolCestStartRows(cestIdx), mPoolCestEndRows(cestIdx))
    sourceSFN = mPoolCestSFN(cestIdx)
    mResolveOneCsetExtractSeconds = mResolveOneCsetExtractSeconds + (MicroTimer_TXSFNCR() - tExtract)

    If DEBUG_TXSFNCR Then
        Debug.Print "TryResolveCset", "cestIdx", cestIdx, "rowCount", rowCount, "sourceSFN", sourceSFN
    End If
    DebugLog "CSET_ANALYZE", "CsetIdx", cestIdx, "SourceSFN", sourceSFN, "CandidateCount", rowCount, "MovesRequired", IIf(rowCount > 0, rowCount - 1, 0), "Rows", JoinLongArray(rows)

    changed = TryPlaceOneMovedRow_NoSourceRetest(rows, rowCount, sourceSFN, movedRowIdx)

    If changed Then
        ResolveOneCset = True
        mPoolCountResolved = mPoolCountResolved + 1

        If DEBUG_TXSFNCR Then
            Debug.Print "CSET resolved", "cestIdx", cestIdx, "movedRowIdx", movedRowIdx, "newSFN", mCurrentSFN(movedRowIdx)
        End If
        DebugLog "CSET_RESOLVED", "CsetIdx", cestIdx, "RowIdx", movedRowIdx, "NewSFN", mCurrentSFN(movedRowIdx), "Reason", "move legal and local group valid"

        tPost = MicroTimer_TXSFNCR()
        BuildPoolFromConflictStart movedRowIdx
        mResolveOneCsetPostSeconds = mResolveOneCsetPostSeconds + (MicroTimer_TXSFNCR() - tPost)
    Else
        mUnresolvedAttemptCount = mUnresolvedAttemptCount + 1
        DebugLog "CSET_UNRESOLVED", "CsetIdx", cestIdx, "SourceSFN", sourceSFN, "Reason", IIf(mLastAttemptReason = vbNullString, "no legal placement found", mLastAttemptReason)
    End If

    mResolveOneCsetSeconds = mResolveOneCsetSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function GetOrderedCestIndexes() As Long()
    Dim idxs() As Long
    Dim used() As Boolean
    Dim i As Long
    Dim outPos As Long
    Dim bestIdx As Long
    Dim bestScore As Double
    Dim score As Double
    Dim haveBest As Boolean

    If mPoolCestCount <= 0 Then
        ReDim idxs(1 To 1)
        idxs(1) = 0
        GetOrderedCestIndexes = idxs
        Exit Function
    End If

    ReDim idxs(1 To mPoolCestCount)
    ReDim used(1 To mPoolCestCount)
    outPos = 1

    Do While outPos <= mPoolCestCount
        haveBest = False
        bestIdx = 0
        bestScore = 0#

        For i = 1 To mPoolCestCount
            If Not used(i) Then
                score = Abs(CDbl(mPoolCestSFN(i)) - mPoolCenter)
                If Not haveBest Then
                    haveBest = True
                    bestIdx = i
                    bestScore = score
                ElseIf score < bestScore Then
                    bestIdx = i
                    bestScore = score
                ElseIf score = bestScore Then
                    If mPoolCestSFN(i) < mPoolCestSFN(bestIdx) Then
                        bestIdx = i
                        bestScore = score
                    End If
                End If
            End If
        Next i

        idxs(outPos) = bestIdx
        used(bestIdx) = True
        outPos = outPos + 1
    Loop

    If DEBUG_TXSFNCR Then
        Dim s As String
        s = "CSET order: "
        For i = LBound(idxs) To UBound(idxs)
            s = s & idxs(i) & IIf(i < UBound(idxs), ",", "")
        Next i
        Debug.Print s & " | poolCenter=" & mPoolCenter
    End If

    GetOrderedCestIndexes = idxs
End Function

Private Function TryPlaceOneMovedRow_NoSourceRetest(ByRef candidateRows() As Long, ByVal candidateCount As Long, ByVal sourceSFN As Long, ByRef movedRowIdx As Long) As Boolean
    Dim i As Long
    Dim rowIdx As Long
    Dim testSFN As Long
    Dim delta As Long
    Dim originalSFN As Long
    Dim maxDec As Long
    Dim maxInc As Long
    Dim maxDelta As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    movedRowIdx = 0

    maxDec = GetNamedLong("maxTX_SFN_est_decrement")
    maxInc = GetNamedLong("maxTX_SFN_est_increment")
    maxDelta = IIf(maxDec > maxInc, maxDec, maxInc)

    If DEBUG_TXSFNCR Then
        Debug.Print "MOVE RANGE", "sourceSFN", sourceSFN, "maxDec", maxDec, "maxInc", maxInc, "maxDelta", maxDelta
    End If

    If maxDelta <= 0 Then
        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    For i = 1 To candidateCount
        rowIdx = candidateRows(i)
        originalSFN = mCurrentSFN(rowIdx)

        For delta = 1 To maxDelta
            If delta <= maxDec Then
                testSFN = sourceSFN - delta
                If DEBUG_TXSFNCR Then Debug.Print "TEST MOVE", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "delta", delta, "testSFN", testSFN, "dir", "-"
                DebugLog "MOVE_ATTEMPT", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Direction", "-", "CandidateOrdinal", i
                If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                    mCurrentSFN(rowIdx) = testSFN
                    If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                        movedRowIdx = rowIdx
                        If DEBUG_TXSFNCR Then Debug.Print "ACCEPT move", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "testSFN", testSFN, "minRx", mRowMinRxTime(rowIdx), "txID", mRowTXID(rowIdx)
                        DebugLog "MOVE_ACCEPTED", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Direction", "-", "Result", mLastAttemptReason
                        TryPlaceOneMovedRow_NoSourceRetest = True
                        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
                        Exit Function
                    End If
                    DebugLog "MOVE_REJECTED", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Reason", mLastAttemptReason
                    mCurrentSFN(rowIdx) = originalSFN
                Else
                    DebugLog "MOVE_REJECTED", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Reason", mLastAttemptReason
                End If
            End If

            If delta <= maxInc Then
                testSFN = sourceSFN + delta
                If DEBUG_TXSFNCR Then Debug.Print "TEST MOVE", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "delta", delta, "testSFN", testSFN, "dir", "+"
                DebugLog "MOVE_ATTEMPT", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Direction", "+", "CandidateOrdinal", i
                If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                    mCurrentSFN(rowIdx) = testSFN
                    If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                        movedRowIdx = rowIdx
                        If DEBUG_TXSFNCR Then Debug.Print "ACCEPT move", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "testSFN", testSFN, "minRx", mRowMinRxTime(rowIdx), "txID", mRowTXID(rowIdx)
                        DebugLog "MOVE_ACCEPTED", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Direction", "+", "Result", mLastAttemptReason
                        TryPlaceOneMovedRow_NoSourceRetest = True
                        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
                        Exit Function
                    End If
                    DebugLog "MOVE_REJECTED", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Reason", mLastAttemptReason
                    mCurrentSFN(rowIdx) = originalSFN
                Else
                    DebugLog "MOVE_REJECTED", "CsetIdx", mCurrentCsetIdx, "RowIdx", rowIdx, "SourceSFN", sourceSFN, "TestSFN", testSFN, "Reason", mLastAttemptReason
                End If
            End If
        Next delta
    Next i

    mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    Dim t As Double
    t = MicroTimer_TXSFNCR()
    mLastAttemptReason = vbNullString

    If rowIdx < 1 Or rowIdx > mFilteredCount Then
        mLastAttemptReason = "row out of range"
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "out of range"
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If testSFN = sourceSFN Then
        mLastAttemptReason = "testSFN equals sourceSFN"
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "testSFN equals sourceSFN", "sfn", sourceSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not IsBitmapSFNAllowed(testSFN) Then
        mLastAttemptReason = "bitmap disallows testSFN"
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "bitmap disallows", "testSFN", testSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not EvaluatePoolBucketExcludingRow(sourceSFN, rowIdx) Then
        mLastAttemptReason = "excluding-source bucket invalid"
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "excluding-source bucket invalid", "sourceSFN", sourceSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not EvaluatePoolBucketWithAddedRow(testSFN, rowIdx) Then
        mLastAttemptReason = "added-row bucket invalid"
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "added-row bucket invalid", "testSFN", testSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    IsOneMovedRowPlacementLegal_NoSourceRetest = True
    mLastAttemptReason = "legal"
    mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function TryForwardEscapeMove_NoSourceRetest(ByVal sourceSFN As Long, ByVal rowIdx As Long) As Boolean
    TryForwardEscapeMove_NoSourceRetest = False
End Function

Private Function ResolveTripleSplitDeterministic(ByRef rows() As Long, ByVal sourceSFN As Long) As Boolean
    Dim i As Long

    For i = LBound(rows) To UBound(rows)
        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN - 1) Then
            mCurrentSFN(rows(i)) = sourceSFN - 1
            ResolveTripleSplitDeterministic = True
            Exit Function
        End If

        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN + 1) Then
            mCurrentSFN(rows(i)) = sourceSFN + 1
            ResolveTripleSplitDeterministic = True
            Exit Function
        End If
    Next i
End Function

Private Function IsPoolForcedSplitFirstMoveLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsPoolForcedSplitFirstMoveLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN)
End Function

Private Function IsPoolSingleRowPlacementLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsPoolSingleRowPlacementLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN)
End Function

Private Function EvaluatePoolBucketExcludingRow(ByVal sfnVal As Long, ByVal excludeRowIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    rows = CollectRowsForSFN(sfnVal, excludeRowIdx, -1, rowCount)
    EvaluatePoolBucketExcludingRow = ValidateBucketRows(rows, rowCount)
    DebugLog "BUCKET_VALIDATE", "Type", "ExcludeRow", "SFN", sfnVal, "RowIdx", excludeRowIdx, "Rows", JoinLongArray(rows), "IsValid", EvaluatePoolBucketExcludingRow
    mBucketExcludeSeconds = mBucketExcludeSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function EvaluatePoolBucketWithAddedRow(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    rows = CollectRowsForSFN(sfnVal, -1, addedRowIdx, rowCount)
    EvaluatePoolBucketWithAddedRow = ValidateBucketRows(rows, rowCount)
    DebugLog "BUCKET_VALIDATE", "Type", "AddRow", "SFN", sfnVal, "RowIdx", addedRowIdx, "Rows", JoinLongArray(rows), "IsValid", EvaluatePoolBucketWithAddedRow
    mBucketAddSeconds = mBucketAddSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsMoveWithinRowBounds(ByVal rowIdx As Long, ByVal testSFN As Long) As Boolean
    Dim prevSFN As Long
    Dim nextSFN As Long

    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If testSFN < 0 Then Exit Function

    prevSFN = testSFN
    nextSFN = testSFN

    If rowIdx > 1 Then prevSFN = mCurrentSFN(rowIdx - 1)
    If rowIdx < mFilteredCount Then nextSFN = mCurrentSFN(rowIdx + 1)

    If rowIdx > 1 And testSFN < prevSFN Then Exit Function
    If rowIdx < mFilteredCount And testSFN > nextSFN Then Exit Function

    IsMoveWithinRowBounds = True
End Function

Private Function IsBitmapSFNAllowed(ByVal testSFN As Long) As Boolean
    Dim bitIdx As Long
    Dim bitChar As String

    If mBitmapLen <= 0 Or LenB(mTxBitmap) = 0 Then
        IsBitmapSFNAllowed = True
        Exit Function
    End If

    bitIdx = ((testSFN Mod mBitmapLen) + mBitmapLen) Mod mBitmapLen + 1

    If bitIdx > Len(mTxBitmap) Then
        IsBitmapSFNAllowed = True
        Exit Function
    End If

    bitChar = Mid$(mTxBitmap, bitIdx, 1)
    IsBitmapSFNAllowed = (bitChar <> "0")
End Function

Private Sub WriteResolvedPoolToOutput()
    Dim i As Long
    Dim rowIdx As Long

    For i = 1 To mPoolCount
        rowIdx = mPoolRows(i)
        If rowIdx >= 1 And rowIdx <= mFilteredCount Then
            CopyRowToOutput rowIdx
            mWritten(rowIdx) = True
        End If
        If (i Mod PROGRESS_STEP_ROWS) = 0 Then UpdateProgressBar i
    Next i
End Sub

Private Sub WriteUnwrittenRowsToOutput()
    Dim r As Long

    For r = 1 To mFilteredCount
        If Not mWritten(r) Then
            CopyRowToOutput r
            mWritten(r) = True
        End If
        If (r Mod PROGRESS_STEP_ROWS) = 0 Then UpdateProgressBar r
    Next r
End Sub

Private Sub RecomputeFinalTXperSFN()
    Dim r As Long
    Dim startR As Long
    Dim endR As Long
    Dim curSFN As Long
    Dim runLen As Long
    Dim i As Long

    If mFilteredCount <= 0 Then Exit Sub

    r = 1
    Do While r <= mFilteredCount
        curSFN = mCurrentSFN(r)
        startR = r

        Do While r <= mFilteredCount
            If mCurrentSFN(r) <> curSFN Then Exit Do
            r = r + 1
        Loop

        endR = r - 1
        runLen = endR - startR + 1

        If mIdxTXperSFN > 0 Then
            For i = startR To endR
                mOutputData(i, mIdxTXperSFN) = runLen
            Next i
        End If
    Loop
End Sub

Private Sub ValidateResolvedRXTimingOnly()
    Dim r As Long
    Dim rxMin As Double

    mRemainingViolations = 0
    For r = 1 To mFilteredCount
        rxMin = mRowMinRxTime(r)
        If rxMin <> NO_RX_TIME Then
            If CDbl(mCurrentSFN(r)) > rxMin Then
                mRemainingViolations = mRemainingViolations + 1
            End If
        End If
        If (r Mod PROGRESS_STEP_ROWS) = 0 Then UpdateProgressBar r
    Next r
End Sub

Private Sub FinalizeOutputVariant(ByRef data As Variant)
    data = mOutputData
End Sub

Private Sub CopyRowToOutput(ByVal rowIdx As Long)
    Dim c As Long
    Dim colCount As Long

    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Sub

    colCount = UBound(mOutputData, 2)
    For c = LBound(mData, 2) To colCount
        mOutputData(rowIdx, c) = mData(rowIdx, c)
    Next c

    mOutputData(rowIdx, mIdxSFNCol) = mCurrentSFN(rowIdx)
End Sub

Private Function BuildSubsetExcludingOne(ByRef rowListIn() As Long, ByVal rowCountIn As Long, ByVal removeRowIdx As Long, ByRef rowListOut() As Long) As Long
    BuildSubsetExcludingOne = 0
End Function

Private Sub Sort3RowsByMinRxTime(ByRef rowA As Long, ByRef rowB As Long, ByRef rowC As Long)
End Sub

Private Sub QuickSortLongs(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
End Sub

Private Sub SortRowIndexByCurrentSFN(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
End Sub

Private Function CompareRowOrder(ByVal rowA As Long, ByVal rowB As Long) As Long
    CompareRowOrder = Sgn(mCurrentSFN(rowA) - mCurrentSFN(rowB))
End Function

Private Sub UpdateStatusBar()
    Application.StatusBar = "TX_SFN conflict resolution running..."
End Sub

Private Sub AddDiag(ByVal eventType As String, ByVal v1 As String, ByVal v2 As String, ByVal v3 As String, ByVal v4 As String, ByVal msg As String)
End Sub

Private Sub HistAddLong(ByRef dictObj As Object, ByVal keyVal As Long)
End Sub

Private Sub DumpHistogram(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, ByVal titleText As String, ByRef dictObj As Object)
End Sub

Private Function SafeDiv(ByVal numerator As Double, ByVal denominator As Double) As Double
    If denominator = 0# Then
        SafeDiv = 0#
    Else
        SafeDiv = numerator / denominator
    End If
End Function

Private Sub WriteDiagnosticLog_TXSFNCR(ByVal totalRows As Long, ByVal calcTime As Double)
End Sub

Private Function MicroTimer_TXSFNCR() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency

    If QueryPerformanceFrequency_TXSFNCR(cyFreq) <> 0 Then
        QueryPerformanceCounter_TXSFNCR cyTicks
        If cyFreq > 0 Then MicroTimer_TXSFNCR = cyTicks / cyFreq
    End If
End Function

Private Sub InitializeOutputBuffer()
    Dim colCount As Long

    If mFilteredCount <= 0 Then Exit Sub

    colCount = UBound(mData, 2)
    ReDim mOutputData(1 To mFilteredCount, 1 To colCount)
    mOutputCount = 0
End Sub

Private Function ExtractPoolRows(ByVal startIdx As Long, ByVal endIdx As Long) As Long()
    Dim rows() As Long
    Dim i As Long
    Dim n As Long

    n = endIdx - startIdx + 1
    ReDim rows(1 To n)

    For i = 1 To n
        rows(i) = mPoolRows(startIdx + i - 1)
    Next i

    ExtractPoolRows = rows
End Function

Private Function CollectRowsForSFN(ByVal sfnVal As Long, ByVal excludeRowIdx As Long, ByVal includeRowIdx As Long, ByRef rowCount As Long) As Long()
    Dim rows() As Long
    Dim r As Long

    ReDim rows(1 To mFilteredCount)
    rowCount = 0

    For r = 1 To mFilteredCount
        If r <> excludeRowIdx Then
            If mCurrentSFN(r) = sfnVal Or r = includeRowIdx Then
                rowCount = rowCount + 1
                rows(rowCount) = r
            End If
        End If
        If (r Mod PROGRESS_STEP_ROWS) = 0 Then UpdateProgressBar r
    Next r

    If rowCount = 0 Then
        ReDim rows(1 To 1)
    ElseIf rowCount < mFilteredCount Then
        ReDim Preserve rows(1 To rowCount)
    End If

    CollectRowsForSFN = rows
End Function

Private Function ValidateBucketRows(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim t As Double

    t = MicroTimer_TXSFNCR()

    If rowCount <= 1 Then
        ValidateBucketRows = True
        mLastAttemptReason = "bucket row count <= 1"
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not IsGroupTXIDUnique(rows, rowCount) Then
        mLastAttemptReason = "TX_ID duplicate in bucket"
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not IsGroupMergedStationsUnique(rows, rowCount) Then
        mLastAttemptReason = "RX station duplicate in bucket"
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    ValidateBucketRows = True
    mLastAttemptReason = "bucket valid"
    mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsGroupTXIDUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim dictTx As Object
    Dim i As Long
    Dim rowIdx As Long
    Dim txIDKey As String

    Set dictTx = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        rowIdx = rows(i)
        txIDKey = CStr(mRowTXID(rowIdx))
        If dictTx.Exists(txIDKey) Then Exit Function
        dictTx(txIDKey) = 1
    Next i

    IsGroupTXIDUnique = True
End Function

Private Function IsGroupMergedStationsUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim dictStations As Object
    Dim i As Long
    Dim rowIdx As Long

    Set dictStations = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        rowIdx = rows(i)
        If Not AddRowStationsToDict(rowIdx, dictStations) Then Exit Function
    Next i

    IsGroupMergedStationsUnique = True
End Function

Private Function AddRowStationsToDict(ByVal rowIdx As Long, ByRef dictStations As Object) As Boolean
    Dim txKey As String
    Dim st As Long
    Dim stationId As Long
    Dim stationKey As String
    Dim stationUB As Long

    txKey = CStr(mRowTXID(rowIdx))
    If dictStations.Exists(txKey) Then Exit Function
    dictStations(txKey) = 1

    stationUB = SafeLongArrayUBound(mRxStationIDs)
    For st = 1 To mActiveRxCount
        If HasRowRxForStation(rowIdx, st) Then
            If stationUB > 0 And st <= stationUB Then
                stationId = mRxStationIDs(st)
            Else
                stationId = st
            End If

            stationKey = CStr(stationId)
            If dictStations.Exists(stationKey) Then Exit Function
            dictStations(stationKey) = 1
        End If
    Next st

    AddRowStationsToDict = True
End Function

Private Function HasRowRxForStation(ByVal rowIdx As Long, ByVal stationOrdinal As Long) As Boolean
    Dim colIdx As Long
    Dim dataUB As Long

    If stationOrdinal < 1 Or stationOrdinal > mActiveRxCount Then Exit Function

    dataUB = SafeLongArrayUBound(mRxDataColIdx)
    If dataUB <= 0 Then Exit Function
    If stationOrdinal > dataUB Then Exit Function

    colIdx = mRxDataColIdx(stationOrdinal)
    If colIdx <= 0 Then Exit Function
    If IsEmpty(mData(rowIdx, colIdx)) Then Exit Function

    HasRowRxForStation = (LenB(Trim$(CStr(mData(rowIdx, colIdx)))) > 0)
End Function

Private Function DoesMovedRowFormValidLocalGroup(ByVal movedRowIdx As Long) As Boolean
    Dim movedSFN As Long

    movedSFN = mCurrentSFN(movedRowIdx)

    If Not EvaluatePoolBucketWithAddedRow(movedSFN, movedRowIdx) Then
        mLastAttemptReason = "moved row local group invalid"
        If DEBUG_TXSFNCR Then Debug.Print "REJECT local-group", "movedRowIdx", movedRowIdx, "movedSFN", movedSFN
        Exit Function
    End If

    mLastAttemptReason = "moved row local group valid"
    DoesMovedRowFormValidLocalGroup = True
End Function

Private Function GetMaxMoveOffset(ByVal sourceSFN As Long) As Long
    Dim lowerRoom As Long
    Dim bitmapLimit As Long
    Dim maxOffset As Long

    lowerRoom = sourceSFN
    bitmapLimit = mBitmapLen
    If bitmapLimit <= 0 Then bitmapLimit = 64

    maxOffset = bitmapLimit
    If lowerRoom < maxOffset Then maxOffset = lowerRoom
    If maxOffset < 1 Then maxOffset = 1

    If DEBUG_TXSFNCR Then
        Debug.Print "GetMaxMoveOffset", "sourceSFN", sourceSFN, "lowerRoom", lowerRoom, "bitmapLimit", bitmapLimit, "maxOffset", maxOffset
    End If

    GetMaxMoveOffset = maxOffset
End Function

Private Function GetRowMinRxTime(ByVal rowIdx As Long) As Double
    Dim st As Long
    Dim colIdx As Long
    Dim rxVal As Double
    Dim dataUB As Long
    Dim hasValue As Boolean

    GetRowMinRxTime = NO_RX_TIME
    dataUB = SafeLongArrayUBound(mRxDataColIdx)
    If dataUB <= 0 Then Exit Function

    For st = 1 To mActiveRxCount
        If st > dataUB Then Exit For
        colIdx = mRxDataColIdx(st)
        If colIdx > 0 Then
            If Not IsEmpty(mData(rowIdx, colIdx)) Then
                If IsNumeric(mData(rowIdx, colIdx)) Then
                    rxVal = CDbl(mData(rowIdx, colIdx))
                    If Not hasValue Then
                        hasValue = True
                        GetRowMinRxTime = rxVal
                    ElseIf rxVal < GetRowMinRxTime Then
                        GetRowMinRxTime = rxVal
                    End If
                End If
            End If
        End If
    Next st
End Function

Private Function SafeLongArrayUBound(ByRef arr() As Long) As Long
    On Error Resume Next
    SafeLongArrayUBound = UBound(arr)
    If Err.Number <> 0 Then
        SafeLongArrayUBound = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function IsAllOnesBitmap(ByVal txBitmap As String) As Boolean
    Dim i As Long

    If Len(txBitmap) = 0 Then Exit Function

    For i = 1 To Len(txBitmap)
        If Mid$(txBitmap, i, 1) <> "1" Then Exit Function
    Next i

    IsAllOnesBitmap = True
End Function

Private Function GetNamedLong(ByVal nameText As String) As Long
    On Error Resume Next
    GetNamedLong = CLng(Evaluate(ThisWorkbook.Names(nameText).RefersTo))
    On Error GoTo 0
End Function

Private Function GetRequiredColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    On Error GoTo MissingColumn
    GetRequiredColumnIndex = lo.ListColumns(columnName).Index
    Exit Function
MissingColumn:
    Err.Clear
    Err.Raise vbObjectError + 2200, "TX_SFNConflictResolution_DEBUG", "Required column not found: " & columnName
End Function

Private Sub BuildRxColumnMaps(ByVal lo As ListObject, ByRef rxDataColIdx() As Long, ByRef rxStationIDs() As Long, ByRef activeRxCount As Long)
    Dim i As Long
    Dim headerText As String
    Dim stationId As Long

    activeRxCount = 0
    For i = 1 To lo.ListColumns.Count
        headerText = UCase$(Trim$(lo.ListColumns(i).Name))
        If Left$(headerText, 6) = "RXTIME" Then
            activeRxCount = activeRxCount + 1
        End If
    Next i

    If activeRxCount <= 0 Then
        ReDim rxDataColIdx(1 To 1)
        ReDim rxStationIDs(1 To 1)
        activeRxCount = 1
        rxDataColIdx(1) = 0
        rxStationIDs(1) = 1
        Exit Sub
    End If

    ReDim rxDataColIdx(1 To activeRxCount)
    ReDim rxStationIDs(1 To activeRxCount)

    activeRxCount = 0
    For i = 1 To lo.ListColumns.Count
        headerText = UCase$(Trim$(lo.ListColumns(i).Name))
        If Left$(headerText, 6) = "RXTIME" Then
            activeRxCount = activeRxCount + 1
            rxDataColIdx(activeRxCount) = i
            stationId = Val(Mid$(headerText, 7))
            If stationId <= 0 Then stationId = activeRxCount
            rxStationIDs(activeRxCount) = stationId
        End If
    Next i
End Sub

Private Function GetConfigText(ByVal nameText As String) As String
    On Error Resume Next
    GetConfigText = CStr(ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range(nameText).Value)
    If LenB(GetConfigText) = 0 Then
        GetConfigText = CStr(Evaluate(ThisWorkbook.Names(nameText).RefersTo))
    End If
    On Error GoTo 0
End Function

Private Function GetConfigLong(ByVal nameText As String) As Long
    On Error Resume Next
    GetConfigLong = CLng(ThisWorkbook.Sheets("Exp Config & Data Proc Params").Range(nameText).Value)
    If GetConfigLong <= 0 Then
        GetConfigLong = CLng(Evaluate(ThisWorkbook.Names(nameText).RefersTo))
    End If
    On Error GoTo 0
End Function

Private Sub EnsureDebugLogSheet()
    On Error Resume Next
    Set mDebugSheet = ThisWorkbook.Worksheets("Conflict Resolution Debug Log")
    On Error GoTo 0

    If mDebugSheet Is Nothing Then
        Set mDebugSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        mDebugSheet.Name = "Conflict Resolution Debug Log"
    End If

    mDebugSheet.Cells.ClearContents
    mDebugSheet.Cells(1, 1).Value = "Seq"
    mDebugSheet.Cells(1, 2).Value = "t(s)"
    mDebugSheet.Cells(1, 3).Value = "Event"
    mDebugSheet.Cells(1, 4).Value = "K1"
    mDebugSheet.Cells(1, 5).Value = "V1"
    mDebugSheet.Cells(1, 6).Value = "K2"
    mDebugSheet.Cells(1, 7).Value = "V2"
    mDebugSheet.Cells(1, 8).Value = "K3"
    mDebugSheet.Cells(1, 9).Value = "V3"
    mDebugSheet.Cells(1, 10).Value = "K4"
    mDebugSheet.Cells(1, 11).Value = "V4"
    mDebugSheet.Cells(1, 12).Value = "K5"
    mDebugSheet.Cells(1, 13).Value = "V5"
    mDebugSheet.Cells(1, 14).Value = "K6"
    mDebugSheet.Cells(1, 15).Value = "V6"
    mDebugSheet.Rows(1).Font.Bold = True

    mDebugRow = 2
    mDebugSeq = 0
End Sub

Private Sub DebugLog(ByVal eventName As String, ParamArray details() As Variant)
    Dim i As Long
    Dim colOut As Long
    Dim detailText As String

    If mDebugSheet Is Nothing Then Exit Sub

    mDebugSeq = mDebugSeq + 1
    mDebugSheet.Cells(mDebugRow, 1).Value = mDebugSeq
    mDebugSheet.Cells(mDebugRow, 2).Value = Format$(MicroTimer_TXSFNCR() - mDebugStartTime, "0.000000")
    mDebugSheet.Cells(mDebugRow, 3).Value = eventName

    colOut = 4
    For i = LBound(details) To UBound(details)
        detailText = VariantToDebugText(details(i))
        mDebugSheet.Cells(mDebugRow, colOut).Value = detailText
        colOut = colOut + 1
        If colOut > 15 Then Exit For
    Next i
    mDebugRow = mDebugRow + 1
End Sub

Private Function VariantToDebugText(ByVal v As Variant) As String
    On Error GoTo Fallback
    If IsObject(v) Then
        VariantToDebugText = TypeName(v)
    ElseIf IsError(v) Then
        VariantToDebugText = "#ERR " & CStr(CLng(v))
    ElseIf IsEmpty(v) Then
        VariantToDebugText = ""
    Else
        VariantToDebugText = CStr(v)
    End If
    Exit Function
Fallback:
    VariantToDebugText = "<unprintable>"
End Function

Private Function JoinLongArray(ByRef arr() As Long) As String
    Dim i As Long

    On Error GoTo NoArray
    For i = LBound(arr) To UBound(arr)
        JoinLongArray = JoinLongArray & CStr(arr(i))
        If i < UBound(arr) Then JoinLongArray = JoinLongArray & ","
    Next i
    Exit Function
NoArray:
    JoinLongArray = ""
End Function

Private Function BuildPoolRowListText() As String
    Dim i As Long
    For i = 1 To mPoolCount
        BuildPoolRowListText = BuildPoolRowListText & mPoolRows(i)
        If i < mPoolCount Then BuildPoolRowListText = BuildPoolRowListText & ","
    Next i
End Function

Private Function BuildRowListFromPoolRange(ByVal startIdx As Long, ByVal endIdx As Long) As String
    Dim i As Long
    For i = startIdx To endIdx
        BuildRowListFromPoolRange = BuildRowListFromPoolRange & mPoolRows(i)
        If i < endIdx Then BuildRowListFromPoolRange = BuildRowListFromPoolRange & ","
    Next i
End Function

Private Sub WriteDebugSummary(ByVal totalLoopCount As Long, ByVal totalFindSeconds As Double, ByVal totalBuildSeconds As Double, ByVal totalResolveSeconds As Double, ByVal totalWriteSeconds As Double, ByVal totalFinalSeconds As Double, ByVal elapsedSeconds As Double)
    DebugLog "SUMMARY_TIMING", "find", Format$(totalFindSeconds, "0.000"), "build", Format$(totalBuildSeconds, "0.000"), "resolve", Format$(totalResolveSeconds, "0.000"), "write", Format$(totalWriteSeconds, "0.000"), "final", Format$(totalFinalSeconds, "0.000")
    DebugLog "SUMMARY_COUNTS", "loops", totalLoopCount, "resolvedCount", mPoolCountResolved, "unresolvedAttempts", mUnresolvedAttemptCount, "remainingViolations", mRemainingViolations, "maxPoolSize", mMaxObservedPoolSize, "elapsedSec", Format$(elapsedSeconds, "0.000")
    mDebugSheet.Columns("A:O").AutoFit
End Sub

Private Sub UpdateProgressBar(ByVal currentRow As Long)
    Application.StatusBar = "TX_SFN conflict resolution DEBUG running... row " & currentRow & " / " & mFilteredCount
End Sub
