Option Explicit

' Version: V1.0.6
Private Const MODULE_VERSION_TXSFNCR As String = "V1.0.6"
Private Const DEBUG_TXSFNCR As Boolean = False

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
Private mPoolLeftRow As Long
Private mPoolRightRow As Long
Private mPoolMinSFN As Long
Private mPoolMaxSFN As Long
Private mPoolCenter As Double
Private mPoolCestStartRows() As Long
Private mPoolCestEndRows() As Long
Private mPoolCestSFN() As Long
Private mPoolCestCount As Long
Private mMaxMoveDec As Long
Private mMaxMoveInc As Long
Private mAllBitmapAllowed As Boolean
Private mRxDataUB As Long
Private mStationGroupIndex() As Long
Private mStationGroupCount As Long
Private mRowHasRx() As Byte
Private mRowHasStationGroup() As Byte
Private mResolveEntirePoolSeconds As Double
Private mResolveOneCsetSeconds As Double
Private mTryPlaceMoveSeconds As Double
Private mLegalCheckSeconds As Double
Private mBucketExcludeSeconds As Double
Private mBucketAddSeconds As Double
Private mValidateBucketSeconds As Double
Private mResolveOneCsetExtractSeconds As Double
Private mResolveOneCsetPostSeconds As Double
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

Public Sub Run_TX_SFNConflictResolution()
    MsgBox "Run_TX_SFNConflictResolution is a wrapper. Call TX_SFNConflictResolution from PickExp.", vbInformation, "TX_SFN Conflict Resolution " & MODULE_VERSION_TXSFNCR
End Sub

Public Sub TX_SFNConflictResolution( _
    ByRef data As Variant, ByVal filteredCount As Long, ByVal idxSFNCol As Long, ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, ByVal idxLEN As Long, ByVal idxTXperSFN As Long, ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, ByVal idxTotLat As Long, ByVal idxGen As Long, ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, ByVal activeRxCount As Long, ByRef dictS2V As Object, ByRef dictVC As Object, _
    ByRef dictA2P As Object, ByRef dictP2R As Object, ByRef dictP2Sigma As Object, ByVal txBitmap As String, _
    ByVal bitmapLen As Long, ByRef elapsedSeconds As Double)

    Dim startTime As Double
    Dim t0 As Double
    Dim conflictStart As Long
    Dim poolMoved As Boolean

    startTime = MicroTimer_TXSFNCR()
    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, txBitmap, bitmapLen
    If Not ValidateInputMonotoneTXSFN() Then Exit Sub
    PrepareRowDerivedData
    InitializeOutputBuffer

    Do
        t0 = MicroTimer_TXSFNCR()
        conflictStart = FindNextConflictStart()
        If conflictStart <= 0 Then Exit Do
        BuildPoolFromConflictStart conflictStart
        poolMoved = ResolveEntirePool()
        WriteResolvedPoolToOutput
        If mPoolCount > 0 Then mScanPos = mPoolRows(mPoolCount) + 1
        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR() - t0)
    Loop

    t0 = MicroTimer_TXSFNCR()
    WriteUnwrittenRowsToOutput
    RecomputeFinalTXperSFN
    ValidateResolvedRXTimingOnly
    FinalizeOutputVariant data
    mFinalizeSeconds = mFinalizeSeconds + (MicroTimer_TXSFNCR() - t0)
    elapsedSeconds = MicroTimer_TXSFNCR() - startTime
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
    mPoolLeftRow = 0
    mPoolRightRow = 0
    mPoolCestCount = 0
    mRxDataUB = SafeLongArrayUBound(mRxDataColIdx)
    InitializeStationGroupIndex
    mMaxMoveDec = GetNamedLong("maxTX_SFN_est_decrement")
    mMaxMoveInc = GetNamedLong("maxTX_SFN_est_increment")
    mAllBitmapAllowed = (mBitmapLen <= 0 Or LenB(mTxBitmap) = 0 Or IsAllOnesBitmap(mTxBitmap))
    If DEBUG_TXSFNCR Then Debug.Print "TX_SFNCR init: filteredCount=" & mFilteredCount & " bitmapLen=" & mBitmapLen
End Sub

Private Function ValidateInputMonotoneTXSFN() As Boolean
    Dim r As Long, prevVal As Long, curVal As Long
    ValidateInputMonotoneTXSFN = True
    If mFilteredCount <= 1 Then Exit Function
    prevVal = CLng(mData(1, mIdxSFNCol))
    For r = 2 To mFilteredCount
        curVal = CLng(mData(r, mIdxSFNCol))
        If curVal < prevVal Then ValidateInputMonotoneTXSFN = False: Exit Function
        prevVal = curVal
    Next r
End Function

Private Sub PrepareRowDerivedData()
    Dim r As Long
    Dim st As Long
    Dim colIdx As Long
    Dim rxVal As Double
    Dim hasValue As Boolean
    Dim cellValue As Variant
    Dim groupIdx As Long

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

    If mActiveRxCount > 0 Then
        ReDim mRowHasRx(1 To mFilteredCount, 1 To mActiveRxCount)
    Else
        ReDim mRowHasRx(1 To 1, 1 To 1)
    End If

    If mStationGroupCount > 0 Then
        ReDim mRowHasStationGroup(1 To mFilteredCount, 1 To mStationGroupCount)
    Else
        ReDim mRowHasStationGroup(1 To 1, 1 To 1)
    End If

    For r = 1 To mFilteredCount
        mInitialSFN(r) = CLng(mData(r, mIdxSFNCol))
        mCurrentSFN(r) = mInitialSFN(r)
        mRowTXID(r) = CLng(mData(r, mIdxTXID))
        mRowTXQTime(r) = CDbl(mData(r, mIdxTXQ))
        mRowNsch(r) = CLng(mData(r, mIdxLEN))
        mRowPduKey(r) = CStr(mData(r, mIdxLEN))
        mRowMinRxTime(r) = NO_RX_TIME
        mRowOriginalIndex(r) = r
        mRowValidInput(r) = True
        If mIdxRxCnt > 0 Then
            mRowRXCount(r) = CLng(mData(r, mIdxRxCnt))
        Else
            mRowRXCount(r) = 0
        End If

        hasValue = False
        For st = 1 To mActiveRxCount
            If st > mRxDataUB Then Exit For
            colIdx = mRxDataColIdx(st)
            If colIdx > 0 Then
                cellValue = mData(r, colIdx)
                If Not IsEmpty(cellValue) Then
                    If LenB(Trim$(CStr(cellValue))) > 0 Then
                        mRowHasRx(r, st) = 1
                        groupIdx = mStationGroupIndex(st)
                        If groupIdx > 0 Then mRowHasStationGroup(r, groupIdx) = 1
                        If IsNumeric(cellValue) Then
                            rxVal = CDbl(cellValue)
                            If Not hasValue Then
                                hasValue = True
                                mRowMinRxTime(r) = rxVal
                            ElseIf rxVal < mRowMinRxTime(r) Then
                                mRowMinRxTime(r) = rxVal
                            End If
                        End If
                    End If
                End If
            End If
        Next st
    Next r
End Sub

Private Function FindNextConflictStart() As Long
    Dim r As Long
    For r = mScanPos To mFilteredCount - 1
        If mCurrentSFN(r) = mCurrentSFN(r + 1) Then FindNextConflictStart = r: Exit Function
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

    mPoolLeftRow = leftRow
    mPoolRightRow = rightRow
    mPoolCount = rightRow - leftRow + 1
    If mPoolCount <= 0 Then Exit Sub

    ReDim mPoolRows(1 To mPoolCount)
    For i = 1 To mPoolCount
        mPoolRows(i) = leftRow + i - 1
    Next i

    mPoolMinSFN = mCurrentSFN(leftRow)
    mPoolMaxSFN = mCurrentSFN(rightRow)
    mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#
    If mPoolCount > mMaxObservedPoolSize Then mMaxObservedPoolSize = mPoolCount
    If DEBUG_TXSFNCR Then Debug.Print "POOL built: startRow=" & startRow & " leftRow=" & leftRow & " rightRow=" & rightRow & " poolCount=" & mPoolCount & " minSFN=" & mPoolMinSFN & " maxSFN=" & mPoolMaxSFN
    BuildPoolCests
End Sub

Private Sub BuildPoolCests()
    Dim i As Long, r As Long, startIdx As Long, curSFN As Long
    If mPoolCount <= 0 Then Exit Sub
    ReDim mPoolCestStartRows(1 To mPoolCount)
    ReDim mPoolCestEndRows(1 To mPoolCount)
    ReDim mPoolCestSFN(1 To mPoolCount)
    mPoolCestCount = 0: i = 1
    Do While i <= mPoolCount
        startIdx = i: r = mPoolRows(i): curSFN = mCurrentSFN(r)
        Do While i < mPoolCount
            If mCurrentSFN(mPoolRows(i + 1)) <> curSFN Then Exit Do
            i = i + 1
        Loop
        mPoolCestCount = mPoolCestCount + 1: mPoolCestStartRows(mPoolCestCount) = startIdx: mPoolCestEndRows(mPoolCestCount) = i: mPoolCestSFN(mPoolCestCount) = curSFN
        If DEBUG_TXSFNCR Then Debug.Print "CEST chunk: idx=" & mPoolCestCount & " rows=" & startIdx & ".." & i & " sfn=" & curSFN & " count=" & (i - startIdx + 1)
        i = i + 1
    Loop
End Sub

Private Function ResolveEntirePool() As Boolean
    Dim orderedCests() As Long
    Dim i As Long
    Dim cestIdx As Long
    Dim changed As Boolean
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    ResolveEntirePool = False
    If mPoolCestCount <= 0 Then Exit Function

    Do
        changed = False
        orderedCests = GetOrderedCestIndexes()
        For i = LBound(orderedCests) To UBound(orderedCests)
            cestIdx = orderedCests(i)
            If cestIdx > 0 Then
                If ResolveOneCset(cestIdx) Then
                    ResolveEntirePool = True
                    changed = True
                    Exit For
                End If
            End If
        Next i
    Loop While changed And mPoolCestCount > 0

    mResolveEntirePoolSeconds = mResolveEntirePoolSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function TryResolveCset(ByRef rows() As Long, ByVal rowCount As Long, ByVal sourceSFN As Long, ByRef movedRowIdx As Long) As Boolean
    Dim analysis As TCsetAnalysis

    analysis = AnalyzeCsetSingleMoves(rows, rowCount)
    TryResolveCset = (analysis.candidateCount > 0 And TryPlaceOneMovedRow_NoSourceRetest(analysis.candidateRows, analysis.candidateCount, sourceSFN, movedRowIdx))
End Function

Private Function AnalyzeCsetSingleMoves(ByRef rows() As Long, ByVal rowCount As Long) As TCsetAnalysis
    Dim a As TCsetAnalysis, i As Long
    a.MovesRequired = IIf(rowCount > 1, rowCount - 1, 0): a.candidateCount = rowCount: ReDim a.candidateRows(1 To rowCount)
    For i = 1 To rowCount: a.candidateRows(i) = rows(i): Next i
    AnalyzeCsetSingleMoves = a
End Function

Private Function ResolveOneCset(ByVal cestIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim movedRowIdx As Long
    Dim sourceSFN As Long
    Dim changed As Boolean
    Dim t As Double
    Dim tExtract As Double
    Dim tPost As Double

    t = MicroTimer_TXSFNCR()
    ResolveOneCset = False
    If cestIdx < 1 Or cestIdx > mPoolCestCount Then Exit Function

    rowCount = mPoolCestEndRows(cestIdx) - mPoolCestStartRows(cestIdx) + 1
    If rowCount <= 1 Then Exit Function

    tExtract = MicroTimer_TXSFNCR()
    rows = ExtractPoolRows(mPoolCestStartRows(cestIdx), mPoolCestEndRows(cestIdx))
    sourceSFN = mPoolCestSFN(cestIdx)
    mResolveOneCsetExtractSeconds = mResolveOneCsetExtractSeconds + (MicroTimer_TXSFNCR() - tExtract)

    changed = TryResolveCset(rows, rowCount, sourceSFN, movedRowIdx)
    If changed Then
        ResolveOneCset = True
        mPoolCountResolved = mPoolCountResolved + 1
        tPost = MicroTimer_TXSFNCR()
        RefreshPoolAfterMove movedRowIdx
        mResolveOneCsetPostSeconds = mResolveOneCsetPostSeconds + (MicroTimer_TXSFNCR() - tPost)
    Else
        mUnresolvedAttemptCount = mUnresolvedAttemptCount + 1
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
    maxDec = mMaxMoveDec
    maxInc = mMaxMoveInc
    If maxDec < 0 Then maxDec = 0
    If maxInc < 0 Then maxInc = 0
    If maxDec = 0 And maxInc = 0 Then
        maxDelta = GetMaxMoveOffset(sourceSFN)
        maxDec = maxDelta
        maxInc = maxDelta
    Else
        maxDelta = IIf(maxDec > maxInc, maxDec, maxInc)
    End If
    If maxDelta <= 0 Then Exit Function

    For i = 1 To candidateCount
        rowIdx = candidateRows(i)
        originalSFN = mCurrentSFN(rowIdx)
        If Not EvaluatePoolBucketExcludingRow(sourceSFN, rowIdx) Then GoTo NextCandidate

        For delta = 1 To maxDelta
            If delta <= maxDec Then
                testSFN = sourceSFN - delta
                If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                    mCurrentSFN(rowIdx) = testSFN
                    If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                        movedRowIdx = rowIdx
                        TryPlaceOneMovedRow_NoSourceRetest = True
                        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
                        Exit Function
                    End If
                    mCurrentSFN(rowIdx) = originalSFN
                End If
            End If

            If delta <= maxInc Then
                testSFN = sourceSFN + delta
                If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                    mCurrentSFN(rowIdx) = testSFN
                    If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                        movedRowIdx = rowIdx
                        TryPlaceOneMovedRow_NoSourceRetest = True
                        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
                        Exit Function
                    End If
                    mCurrentSFN(rowIdx) = originalSFN
                End If
            End If
        Next delta
NextCandidate:
    Next i

    mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If testSFN = sourceSFN Then Exit Function
    If Not IsMoveWithinRowBounds(rowIdx, testSFN) Then Exit Function
    If Not IsBitmapSFNAllowed(testSFN) Then Exit Function
    If Not EvaluatePoolBucketWithAddedRow(testSFN, rowIdx) Then Exit Function
    IsOneMovedRowPlacementLegal_NoSourceRetest = True
    mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function TryForwardEscapeMove_NoSourceRetest(ByVal sourceSFN As Long, ByVal rowIdx As Long) As Boolean: TryForwardEscapeMove_NoSourceRetest = False: End Function
Private Function ResolveTripleSplitDeterministic(ByRef rows() As Long, ByVal sourceSFN As Long) As Boolean
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN - 1) Then mCurrentSFN(rows(i)) = sourceSFN - 1: ResolveTripleSplitDeterministic = True: Exit Function
        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN + 1) Then mCurrentSFN(rows(i)) = sourceSFN + 1: ResolveTripleSplitDeterministic = True: Exit Function
    Next i
End Function
Private Function IsPoolForcedSplitFirstMoveLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean: IsPoolForcedSplitFirstMoveLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN): End Function
Private Function IsPoolSingleRowPlacementLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean: IsPoolSingleRowPlacementLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN): End Function
Private Function EvaluatePoolBucketExcludingRow(ByVal sfnVal As Long, ByVal excludeRowIdx As Long) As Boolean
    Dim leftRow As Long
    Dim rightRow As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    FindBucketBoundsForSource sfnVal, excludeRowIdx, leftRow, rightRow
    EvaluatePoolBucketExcludingRow = ValidateBucketWindow(leftRow, rightRow)
    mBucketExcludeSeconds = mBucketExcludeSeconds + (MicroTimer_TXSFNCR() - t)
End Function
Private Function EvaluatePoolBucketWithAddedRow(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean
    Dim leftRow As Long
    Dim rightRow As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    FindBucketBoundsForTarget sfnVal, addedRowIdx, leftRow, rightRow
    If leftRow > rightRow Then
        EvaluatePoolBucketWithAddedRow = True
    Else
        EvaluatePoolBucketWithAddedRow = ValidateBucketWindowWithCandidate(leftRow, rightRow, addedRowIdx)
    End If
    mBucketAddSeconds = mBucketAddSeconds + (MicroTimer_TXSFNCR() - t)
End Function
Private Function IsMoveWithinRowBounds(ByVal rowIdx As Long, ByVal testSFN As Long) As Boolean
    Dim prevSFN As Long, nextSFN As Long
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
    If mAllBitmapAllowed Then
        IsBitmapSFNAllowed = True
        Exit Function
    End If
    bitIdx = ((testSFN Mod mBitmapLen) + mBitmapLen) Mod mBitmapLen + 1
    If bitIdx > Len(mTxBitmap) Then
        IsBitmapSFNAllowed = True
        Exit Function
    End If
    IsBitmapSFNAllowed = (Mid$(mTxBitmap, bitIdx, 1) <> "0")
End Function

Private Sub WriteResolvedPoolToOutput()
    Dim i As Long, rowIdx As Long
    For i = 1 To mPoolCount
        rowIdx = mPoolRows(i)
        If rowIdx >= 1 And rowIdx <= mFilteredCount Then CopyRowToOutput rowIdx: mWritten(rowIdx) = True
    Next i
End Sub

Private Sub WriteUnwrittenRowsToOutput()
    Dim r As Long
    For r = 1 To mFilteredCount
        If Not mWritten(r) Then CopyRowToOutput r: mWritten(r) = True
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
    Dim r As Long, rxMin As Double
    mRemainingViolations = 0
    For r = 1 To mFilteredCount
        rxMin = mRowMinRxTime(r)
        If rxMin <> NO_RX_TIME Then
            If CDbl(mCurrentSFN(r)) > rxMin Then mRemainingViolations = mRemainingViolations + 1
        End If
    Next r
End Sub

Private Sub FinalizeOutputVariant(ByRef data As Variant): data = mOutputData: End Sub
Private Sub CopyRowToOutput(ByVal rowIdx As Long)
    Dim c As Long
    Dim colCount As Long
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Sub
    colCount = UBound(mOutputData, 2)
    For c = LBound(mData, 2) To colCount
        mOutputData(rowIdx, c) = mData(rowIdx, c)
    Next c
    mOutputData(rowIdx, mIdxSFNCol) = mCurrentSFN(rowIdx)
    If mIdxLEN > 0 Then mOutputData(rowIdx, mIdxLEN) = mRowNsch(rowIdx)
End Sub

Private Function BuildSubsetExcludingOne(ByRef rowListIn() As Long, ByVal rowCountIn As Long, ByVal removeRowIdx As Long, ByRef rowListOut() As Long) As Long: BuildSubsetExcludingOne = 0: End Function
Private Sub Sort3RowsByMinRxTime(ByRef rowA As Long, ByRef rowB As Long, ByRef rowC As Long): End Sub
Private Sub QuickSortLongs(ByRef arr() As Long, ByVal first As Long, ByVal last As Long): End Sub
Private Sub SortRowIndexByCurrentSFN(ByRef arr() As Long, ByVal first As Long, ByVal last As Long): End Sub
Private Function CompareRowOrder(ByVal rowA As Long, ByVal rowB As Long) As Long: CompareRowOrder = Sgn(mCurrentSFN(rowA) - mCurrentSFN(rowB)): End Function
Private Sub UpdateStatusBar(): Application.StatusBar = "TX_SFN conflict resolution running...": End Sub
Private Sub AddDiag(ByVal eventType As String, ByVal v1 As String, ByVal v2 As String, ByVal v3 As String, ByVal v4 As String, ByVal msg As String): End Sub
Private Sub HistAddLong(ByRef dictObj As Object, ByVal keyVal As Long): End Sub
Private Sub DumpHistogram(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, ByVal titleText As String, ByRef dictObj As Object): End Sub
Private Function SafeDiv(ByVal numerator As Double, ByVal denominator As Double) As Double
    If denominator = 0# Then
        SafeDiv = 0#
    Else
        SafeDiv = numerator / denominator
    End If
End Function
Private Sub WriteDiagnosticLog_TXSFNCR(ByVal totalRows As Long, ByVal calcTime As Double): End Sub
Private Function MicroTimer_TXSFNCR() As Double
    Dim cyTicks As Currency, cyFreq As Currency
    If QueryPerformanceFrequency_TXSFNCR(cyFreq) <> 0 Then QueryPerformanceCounter_TXSFNCR cyTicks: If cyFreq > 0 Then MicroTimer_TXSFNCR = cyTicks / cyFreq
End Function
Private Sub InitializeOutputBuffer()
    Dim colCount As Long
    If mFilteredCount <= 0 Then Exit Sub
    colCount = UBound(mData, 2)
    ReDim mOutputData(1 To mFilteredCount, 1 To colCount)
    mOutputCount = 0
End Sub
Private Function ExtractPoolRows(ByVal startIdx As Long, ByVal endIdx As Long) As Long()
    Dim rows() As Long, i As Long, n As Long
    n = endIdx - startIdx + 1
    ReDim rows(1 To n)
    For i = 1 To n: rows(i) = mPoolRows(startIdx + i - 1): Next i
    ExtractPoolRows = rows
End Function

Private Function CollectRowsForSFN(ByVal sfnVal As Long, ByVal excludeRowIdx As Long, ByVal includeRowIdx As Long, ByRef rowCount As Long) As Long()
    Dim rows() As Long, r As Long
    ReDim rows(1 To mFilteredCount)
    rowCount = 0
    For r = 1 To mFilteredCount
        If r <> excludeRowIdx Then
            If mCurrentSFN(r) = sfnVal Or r = includeRowIdx Then
                rowCount = rowCount + 1
                rows(rowCount) = r
            End If
        End If
    Next r
    If rowCount = 0 Then
        ReDim rows(1 To 1)
    ElseIf rowCount < mFilteredCount Then
        ReDim Preserve rows(1 To rowCount)
    End If
    CollectRowsForSFN = rows
End Function

Private Function ValidateBucketRows(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    If rowCount <= 1 Then
        ValidateBucketRows = True
    Else
        For i = 1 To rowCount - 1
            For j = i + 1 To rowCount
                If RowsConflict(rows(i), rows(j)) Then
                    mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
                    Exit Function
                End If
            Next j
        Next i
        ValidateBucketRows = True
    End If
    mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsGroupTXIDUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim i As Long
    Dim j As Long
    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If mRowTXID(rows(i)) = mRowTXID(rows(j)) Then Exit Function
        Next j
    Next i
    IsGroupTXIDUnique = True
End Function

Private Function IsGroupMergedStationsUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim i As Long
    Dim j As Long
    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If RowsShareStationGroup(rows(i), rows(j)) Then Exit Function
        Next j
    Next i
    IsGroupMergedStationsUnique = True
End Function

Private Function AddRowStationsToDict(ByVal rowIdx As Long, ByRef dictStations As Object) As Boolean
    Dim txKey As String, st As Long, stationId As Long, stationKey As String, stationUB As Long
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
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If stationOrdinal < 1 Or stationOrdinal > mActiveRxCount Then Exit Function
    HasRowRxForStation = (mRowHasRx(rowIdx, stationOrdinal) <> 0)
End Function

Private Function DoesMovedRowFormValidLocalGroup(ByVal movedRowIdx As Long) As Boolean
    If DEBUG_TXSFNCR Then
        DoesMovedRowFormValidLocalGroup = EvaluateCurrentBucketAtRow(movedRowIdx)
    Else
        DoesMovedRowFormValidLocalGroup = True
    End If
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
    GetMaxMoveOffset = maxOffset
End Function

Private Function GetRowMinRxTime(ByVal rowIdx As Long) As Double
    If rowIdx < 1 Or rowIdx > mFilteredCount Then
        GetRowMinRxTime = NO_RX_TIME
    Else
        GetRowMinRxTime = mRowMinRxTime(rowIdx)
    End If
End Function

Private Sub InitializeStationGroupIndex()
    Dim dictStations As Object
    Dim st As Long
    Dim stationId As Long
    Dim stationKey As String

    If mActiveRxCount <= 0 Then
        ReDim mStationGroupIndex(1 To 1)
        mStationGroupCount = 0
        Exit Sub
    End If

    ReDim mStationGroupIndex(1 To mActiveRxCount)
    Set dictStations = CreateObject("Scripting.Dictionary")

    For st = 1 To mActiveRxCount
        If SafeLongArrayUBound(mRxStationIDs) > 0 And st <= SafeLongArrayUBound(mRxStationIDs) Then
            stationId = mRxStationIDs(st)
        Else
            stationId = st
        End If
        stationKey = CStr(stationId)
        If Not dictStations.Exists(stationKey) Then
            mStationGroupCount = mStationGroupCount + 1
            dictStations(stationKey) = mStationGroupCount
        End If
        mStationGroupIndex(st) = CLng(dictStations(stationKey))
    Next st
End Sub

Private Sub RefreshPoolAfterMove(ByVal movedRowIdx As Long)
    Dim leftRow As Long
    Dim rightRow As Long
    Dim i As Long

    leftRow = mPoolLeftRow
    rightRow = mPoolRightRow
    If leftRow < 1 Or rightRow < leftRow Then Exit Sub

    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = mCurrentSFN(leftRow)
        leftRow = leftRow - 1
    Loop
    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = mCurrentSFN(rightRow)
        rightRow = rightRow + 1
    Loop

    If leftRow <> mPoolLeftRow Or rightRow <> mPoolRightRow Then
        mPoolLeftRow = leftRow
        mPoolRightRow = rightRow
        mPoolCount = rightRow - leftRow + 1
        ReDim mPoolRows(1 To mPoolCount)
        For i = 1 To mPoolCount
            mPoolRows(i) = leftRow + i - 1
        Next i
        mPoolMinSFN = mCurrentSFN(leftRow)
        mPoolMaxSFN = mCurrentSFN(rightRow)
        mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#
    End If

    BuildPoolCests
End Sub

Private Sub FindBucketBoundsForSource(ByVal sfnVal As Long, ByVal excludeRowIdx As Long, ByRef leftRow As Long, ByRef rightRow As Long)
    leftRow = excludeRowIdx - 1
    Do While leftRow >= 1
        If mCurrentSFN(leftRow) <> sfnVal Then Exit Do
        leftRow = leftRow - 1
    Loop
    leftRow = leftRow + 1

    rightRow = excludeRowIdx + 1
    Do While rightRow <= mFilteredCount
        If mCurrentSFN(rightRow) <> sfnVal Then Exit Do
        rightRow = rightRow + 1
    Loop
    rightRow = rightRow - 1
End Sub

Private Sub FindBucketBoundsForTarget(ByVal sfnVal As Long, ByVal addedRowIdx As Long, ByRef leftRow As Long, ByRef rightRow As Long)
    leftRow = addedRowIdx - 1
    Do While leftRow >= 1
        If mCurrentSFN(leftRow) <> sfnVal Then Exit Do
        leftRow = leftRow - 1
    Loop
    leftRow = leftRow + 1

    rightRow = addedRowIdx + 1
    Do While rightRow <= mFilteredCount
        If mCurrentSFN(rightRow) <> sfnVal Then Exit Do
        rightRow = rightRow + 1
    Loop
    rightRow = rightRow - 1
End Sub

Private Function EvaluateCurrentBucketAtRow(ByVal rowIdx As Long) As Boolean
    Dim leftRow As Long
    Dim rightRow As Long
    Dim sfnVal As Long

    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    sfnVal = mCurrentSFN(rowIdx)
    leftRow = rowIdx
    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = sfnVal
        leftRow = leftRow - 1
    Loop
    rightRow = rowIdx
    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = sfnVal
        rightRow = rightRow + 1
    Loop
    EvaluateCurrentBucketAtRow = ValidateBucketWindow(leftRow, rightRow)
End Function

Private Function ValidateBucketWindow(ByVal leftRow As Long, ByVal rightRow As Long) As Boolean
    Dim rowA As Long
    Dim rowB As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    If leftRow <= 0 Or rightRow <= 0 Or leftRow >= rightRow Then
        ValidateBucketWindow = True
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    For rowA = leftRow To rightRow - 1
        For rowB = rowA + 1 To rightRow
            If RowsConflict(rowA, rowB) Then
                mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
                Exit Function
            End If
        Next rowB
    Next rowA

    ValidateBucketWindow = True
    mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function ValidateBucketWindowWithCandidate(ByVal leftRow As Long, ByVal rightRow As Long, ByVal candidateRow As Long) As Boolean
    Dim rowA As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    For rowA = leftRow To rightRow
        If rowA <> candidateRow Then
            If RowsConflict(candidateRow, rowA) Then
                mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
                Exit Function
            End If
        End If
    Next rowA

    If DEBUG_TXSFNCR Then
        ValidateBucketWindowWithCandidate = ValidateBucketWindow(leftRow, rightRow)
    Else
        ValidateBucketWindowWithCandidate = True
    End If
    mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function RowsConflict(ByVal rowA As Long, ByVal rowB As Long) As Boolean
    If rowA = rowB Then Exit Function
    If mRowTXID(rowA) = mRowTXID(rowB) Then
        RowsConflict = True
        Exit Function
    End If
    RowsConflict = RowsShareStationGroup(rowA, rowB)
End Function

Private Function RowsShareStationGroup(ByVal rowA As Long, ByVal rowB As Long) As Boolean
    Dim groupIdx As Long
    If mStationGroupCount <= 0 Then Exit Function
    For groupIdx = 1 To mStationGroupCount
        If mRowHasStationGroup(rowA, groupIdx) <> 0 Then
            If mRowHasStationGroup(rowB, groupIdx) <> 0 Then
                RowsShareStationGroup = True
                Exit Function
            End If
        End If
    Next groupIdx
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

Private Function SafeLongArrayUBound(ByRef arr() As Long) As Long
    On Error Resume Next
    SafeLongArrayUBound = UBound(arr)
    If Err.Number <> 0 Then
        SafeLongArrayUBound = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function
