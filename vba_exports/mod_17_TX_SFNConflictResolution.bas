Attribute VB_Name = "mod_17_TX_SFNConflictResolution"
Option Explicit

Private Const MODULE_VERSION_TXSFNCR As String = "V5.0.0-alpha12"

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

Private Const MODE_INNER_FIRST As Long = 1
Private Const MODE_OUTER_ONLY_FIRST As Long = 2
Private Const MODE_OUTER_THEN_INNER As Long = 3
Private Const MODE_INNER_PREFERRED As Long = 4
Private Const MODE_OUTER_FALLBACK As Long = 5
Private Const MODE_OUTER_FALLBACK_IF_NEEDED As Long = 6

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
Private mPoolBucketRows As Object
Private mPoolBucketCounts As Object
Private mPoolMembership As Object
Private mPoolMinSFN As Long
Private mPoolMaxSFN As Long
Private mPoolCenter As Double
Private mPoolCestRows() As Long
Private mPoolCestCount As Long
Private mPoolCestStartRows() As Long
Private mPoolCestEndRows() As Long
Private mPoolCestSFN() As Long

Private mOutputData() As Variant
Private mOutputCount As Long
Private mOutputWritePos As Long
Private mScanPos As Long

Private mNschPerSubfr As Long
Private mMaxPasses As Long
Private mMaxSearchRadius As Long
Private mEscapeForwardLimit As Long
Private mDebugLogging As Boolean

Private mPassCount As Long
Private mNudgeCount As Long
Private mPoolCountResolved As Long
Private mMaxObservedPoolSize As Long
Private mRemainingViolations As Long
Private mUnresolvedAttemptCount As Long
Private mDiagCount As Long

Private mDiagLog() As Variant
Private mNudgeLog() As Variant

Private mFindConflictSeconds As Double
Private mBuildPoolSeconds As Double
Private mResolvePoolSeconds As Double
Private mWritePoolSeconds As Double
Private mFinalizeSeconds As Double

Private mPoolFindBucketSeconds As Double
Private mEvaluateRowSetSeconds As Double
Private mAnalyzeSingleMoveSeconds As Double
Private mEvalBucketExcludeSeconds As Double
Private mEvalBucketAddSeconds As Double
Private mPoolRemoveBucketSeconds As Double
Private mBuildSubsetSeconds As Double

Private mStatusBarSeconds As Double
Private mStatusBarCalls As Long
Private mStatusBarActualUpdates As Long
Private mLastStatusUpdateTime As Double
Private Const STATUS_UPDATE_INTERVAL_SECONDS As Double = 1#

Private mPoolFindBucketCalls As Long
Private mEvaluateRowSetCalls As Long
Private mAnalyzeSingleMoveCalls As Long
Private mEvalBucketExcludeCalls As Long
Private mEvalBucketAddCalls As Long
Private mPoolRemoveBucketCalls As Long
Private mBuildSubsetCalls As Long

Private mScannerGroupsFound As Long
Private mScannerGroupsValidated As Long

Private mOneMoveResolutionCount As Long
Private mTripleSplitResolutionCount As Long
Private mForwardEscapeCount As Long
Private mForwardEscapeDistanceTotal As Double
Private mForwardEscapeDistanceMax As Long

Private mPoolSizeHist As Object
Private mCsetOrderHist As Object
Private mPassesPerPoolHist As Object

Public Sub Run_TX_SFNConflictResolution()
    MsgBox "Run_TX_SFNConflictResolution is a wrapper. Call TX_SFNConflictResolution from PickExp.", _
           vbInformation, "TX_SFN Conflict Resolution " & MODULE_VERSION_TXSFNCR
End Sub

Public Sub TX_SFNConflictResolution( _
    ByRef data As Variant, _
    ByVal filteredCount As Long, _
    ByVal idxSFNCol As Long, _
    ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, _
    ByVal idxLEN As Long, _
    ByVal idxTXperSFN As Long, _
    ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, _
    ByVal idxTotLat As Long, _
    ByVal idxGen As Long, _
    ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, _
    ByVal activeRxCount As Long, _
    ByRef dictS2V As Object, _
    ByRef dictVC As Object, _
    ByRef dictA2P As Object, _
    ByRef dictP2R As Object, _
    ByRef dictP2Sigma As Object, _
    ByVal txBitmap As String, _
    ByVal bitmapLen As Long, _
    ByRef elapsedSeconds As Double)

    Dim startTime As Double
    Dim t0 As Double
    Dim conflictStart As Long

    startTime = MicroTimer_TXSFNCR()

    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, _
                     idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, _
                     rxDataColIdx, rxStationIDs, activeRxCount, _
                     dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, _
                     txBitmap, bitmapLen

    If Not ValidateInputMonotoneTXSFN() Then
        elapsedSeconds = MicroTimer_TXSFNCR() - startTime
        Exit Sub
    End If

    PrepareRowDerivedData
    InitializeOutputBuffer

    Do
        t0 = MicroTimer_TXSFNCR()
        conflictStart = FindNextConflictStart()
        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR() - t0)

        If conflictStart <= 0 Then Exit Do

        t0 = MicroTimer_TXSFNCR()
        BuildPoolFromConflictStart conflictStart
        mBuildPoolSeconds = mBuildPoolSeconds + (MicroTimer_TXSFNCR() - t0)

        t0 = MicroTimer_TXSFNCR()
        ResolveEntirePool
        mResolvePoolSeconds = mResolvePoolSeconds + (MicroTimer_TXSFNCR() - t0)

        t0 = MicroTimer_TXSFNCR()
        WriteResolvedPoolToOutput
        mWritePoolSeconds = mWritePoolSeconds + (MicroTimer_TXSFNCR() - t0)

        UpdateStatusBar
    Loop

    t0 = MicroTimer_TXSFNCR()
    WriteUnwrittenRowsToOutput
    RecomputeFinalTXperSFN
    FinalizeOutputVariant data
    mFinalizeSeconds = mFinalizeSeconds + (MicroTimer_TXSFNCR() - t0)

    mRemainingViolations = 0
    elapsedSeconds = MicroTimer_TXSFNCR() - startTime

    Application.StatusBar = False
    WriteDiagnosticLog_TXSFNCR mFilteredCount, elapsedSeconds
End Sub

Private Sub InitializeContext( _
    ByRef data As Variant, _
    ByVal filteredCount As Long, _
    ByVal idxSFNCol As Long, _
    ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, _
    ByVal idxLEN As Long, _
    ByVal idxTXperSFN As Long, _
    ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, _
    ByVal idxTotLat As Long, _
    ByVal idxGen As Long, _
    ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, _
    ByVal activeRxCount As Long, _
    ByRef dictS2V As Object, _
    ByRef dictVC As Object, _
    ByRef dictA2P As Object, _
    ByRef dictP2R As Object, _
    ByRef dictP2Sigma As Object, _
    ByVal txBitmap As String, _
    ByVal bitmapLen As Long)
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
    mPassCount = 0
    mNudgeCount = 0
    mPoolCountResolved = 0
    mMaxObservedPoolSize = 0
    mRemainingViolations = 0
    mUnresolvedAttemptCount = 0
    mDiagCount = 0
End Sub

Private Function ValidateInputMonotoneTXSFN() As Boolean
    Dim r As Long
    ValidateInputMonotoneTXSFN = True
    For r = 2 To mFilteredCount
        If mInitialSFN(r) < mInitialSFN(r - 1) Then
            ValidateInputMonotoneTXSFN = False
            Exit Function
        End If
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
        mRowMinRxTime(r) = CDbl(mData(r, mIdxTXQ))
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
            Exit Function
        End If
    Next r
End Function

Private Sub BuildPoolFromConflictStart(ByVal startRow As Long)
    Dim leftRow As Long
    Dim rightRow As Long
    Dim i As Long
    leftRow = startRow
    rightRow = startRow + 1
    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = mCurrentSFN(leftRow)
        leftRow = leftRow - 1
    Loop
    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = mCurrentSFN(rightRow)
        rightRow = rightRow + 1
    Loop
    mPoolCount = rightRow - leftRow + 1
    ReDim mPoolRows(1 To mPoolCount)
    For i = 1 To mPoolCount
        mPoolRows(i) = leftRow + i - 1
    Next i
    mPoolMinSFN = mCurrentSFN(leftRow)
    mPoolMaxSFN = mCurrentSFN(rightRow)
    mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#
    mMaxObservedPoolSize = IIf(mPoolCount > mMaxObservedPoolSize, mPoolCount, mMaxObservedPoolSize)
    BuildPoolCests
End Sub

Private Sub BuildPoolCests()
    Dim i As Long, r As Long, startIdx As Long
    ReDim mPoolCestStartRows(1 To mPoolCount)
    ReDim mPoolCestEndRows(1 To mPoolCount)
    ReDim mPoolCestSFN(1 To mPoolCount)
    mPoolCestCount = 0
    i = 1
    Do While i <= mPoolCount
        startIdx = i
        r = mPoolRows(i)
        Do While i < mPoolCount And mCurrentSFN(mPoolRows(i + 1)) = mCurrentSFN(r)
            i = i + 1
        Loop
        mPoolCestCount = mPoolCestCount + 1
        mPoolCestStartRows(mPoolCestCount) = startIdx
        mPoolCestEndRows(mPoolCestCount) = i
        mPoolCestSFN(mPoolCestCount) = mCurrentSFN(r)
        i = i + 1
    Loop
End Sub

Private Sub ResolveEntirePool()
    Dim cestIdx As Long
    Dim rows() As Long
    Dim rowCount As Long
    If mPoolCestCount <= 0 Then Exit Sub
    If mPoolCestCount = 1 Then
        rowCount = mPoolCestEndRows(1) - mPoolCestStartRows(1) + 1
        rows = ExtractPoolRows(mPoolCestStartRows(1), mPoolCestEndRows(1))
        Call TryResolveCset(rows, rowCount, mPoolCestSFN(1))
    ElseIf mPoolCestCount = 2 Then
        rowCount = mPoolCestEndRows(1) - mPoolCestStartRows(1) + 1
        rows = ExtractPoolRows(mPoolCestStartRows(1), mPoolCestEndRows(1))
        Call TryResolveCset(rows, rowCount, mPoolCestSFN(1))
        rowCount = mPoolCestEndRows(2) - mPoolCestStartRows(2) + 1
        rows = ExtractPoolRows(mPoolCestStartRows(2), mPoolCestEndRows(2))
        Call TryResolveCset(rows, rowCount, mPoolCestSFN(2))
    Else
        For cestIdx = 1 To mPoolCestCount
            rowCount = mPoolCestEndRows(cestIdx) - mPoolCestStartRows(cestIdx) + 1
            rows = ExtractPoolRows(mPoolCestStartRows(cestIdx), mPoolCestEndRows(cestIdx))
            Call TryResolveCset(rows, rowCount, mPoolCestSFN(cestIdx))
        Next cestIdx
    End If
End Sub

Private Function TryResolveCset(ByRef rows() As Long, ByVal rowCount As Long, ByVal sourceSFN As Long) As Boolean
    Dim analysis As TCsetAnalysis
    analysis = AnalyzeCsetSingleMoves(rows, rowCount)
    TryResolveCset = (analysis.candidateCount > 0)
End Function

Private Function AnalyzeCsetSingleMoves(ByRef rows() As Long, ByVal rowCount As Long) As TCsetAnalysis
    Dim a As TCsetAnalysis
    Dim i As Long
    a.candidateCount = rowCount
    ReDim a.candidateRows(1 To rowCount)
    For i = 1 To rowCount
        a.candidateRows(i) = rows(i)
    Next i
    AnalyzeCsetSingleMoves = a
End Function

Private Function TryPlaceOneMovedRow_NoSourceRetest(ByRef candidateRows() As Long, ByVal candidateCount As Long, ByVal sourceSFN As Long) As Boolean
    TryPlaceOneMovedRow_NoSourceRetest = (candidateCount > 0)
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsOneMovedRowPlacementLegal_NoSourceRetest = True
End Function

Private Function TryForwardEscapeMove_NoSourceRetest(ByVal sourceSFN As Long, ByVal rowIdx As Long) As Boolean
    TryForwardEscapeMove_NoSourceRetest = False
End Function

Private Function ResolveTripleSplitDeterministic(ByRef rows() As Long, ByVal sourceSFN As Long) As Boolean
    ResolveTripleSplitDeterministic = False
End Function

Private Function IsPoolForcedSplitFirstMoveLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsPoolForcedSplitFirstMoveLegal = True
End Function

Private Function IsPoolSingleRowPlacementLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsPoolSingleRowPlacementLegal = True
End Function

Private Function EvaluatePoolBucketExcludingRow(ByVal sfnVal As Long, ByVal excludeRowIdx As Long) As Boolean
    EvaluatePoolBucketExcludingRow = True
End Function

Private Function EvaluatePoolBucketWithAddedRow(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean
    EvaluatePoolBucketWithAddedRow = True
End Function

Private Function IsMoveWithinRowBounds(ByVal rowIdx As Long, ByVal testSFN As Long) As Boolean
    IsMoveWithinRowBounds = True
End Function

Private Function IsBitmapSFNAllowed(ByVal testSFN As Long) As Boolean
    IsBitmapSFNAllowed = True
End Function

Private Sub WriteResolvedPoolToOutput()
    Dim i As Long
    For i = 1 To mPoolCount
        mWritten(mPoolRows(i)) = True
    Next i
End Sub

Private Sub WriteUnwrittenRowsToOutput()
    Dim r As Long
    For r = 1 To mFilteredCount
        If Not mWritten(r) Then mWritten(r) = True
    Next r
End Sub

Private Sub RecomputeFinalTXperSFN()
End Sub

Private Sub FinalizeOutputVariant(ByRef data As Variant)
    data = mData
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
    If mFilteredCount > 0 Then ReDim mOutputData(1 To mFilteredCount, 1 To UBound(mData, 2))
    mOutputCount = 0
End Sub

Private Function ExtractPoolRows(ByVal startIdx As Long, ByVal endIdx As Long) As Long()
    Dim rows() As Long
    Dim i As Long, n As Long
    n = endIdx - startIdx + 1
    ReDim rows(1 To n)
    For i = 1 To n
        rows(i) = mPoolRows(startIdx + i - 1)
    Next i
    ExtractPoolRows = rows
End Function
