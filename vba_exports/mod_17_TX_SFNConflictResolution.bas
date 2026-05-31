Attribute VB_Name = "mod_17_TX_SFNConflictResolution"
Option Explicit

' Version: V1.0.1
Private Const MODULE_VERSION_TXSFNCR As String = "V1.0.1"
Private Const DEBUG_TXSFNCR As Boolean = True

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
Private mMaxTXSFNDecrement As Long
Private mMaxTXSFNIncrement As Long
Private mOutputRowSource() As Long

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

    Dim startTime As Double: startTime = MicroTimer_TXSFNCR()
    Dim t0 As Double
    Dim conflictStart As Long
    Dim poolMoved As Boolean

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
    mData = data: mFilteredCount = filteredCount: mIdxSFNCol = idxSFNCol: mIdxTXID = idxTXID: mIdxTXQ = idxTXQ: mIdxLEN = idxLEN: mIdxTXperSFN = idxTXperSFN: mIdxRxCnt = idxRxCnt: mIdxAvg = idxAvg: mIdxTotLat = idxTotLat: mIdxGen = idxGen
    mRxDataColIdx = rxDataColIdx: mRxStationIDs = rxStationIDs: mActiveRxCount = activeRxCount
    Set mDictS2V = dictS2V: Set mDictVC = dictVC: Set mDictA2P = dictA2P: Set mDictP2R = dictP2R: Set mDictP2Sigma = dictP2Sigma
    mTxBitmap = txBitmap: mBitmapLen = bitmapLen: mScanPos = 1: mOutputWritePos = 1: mPoolCount = 0: mPoolCestCount = 0
    LoadDirectionalMoveBounds
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
        mInitialSFN(r) = CLng(mData(r, mIdxSFNCol)): mCurrentSFN(r) = mInitialSFN(r): mRowTXID(r) = CLng(mData(r, mIdxTXID)): mRowTXQTime(r) = CDbl(mData(r, mIdxTXQ)): mRowNsch(r) = CLng(mData(r, mIdxLEN)): mRowPduKey(r) = CStr(mData(r, mIdxLEN)): mRowMinRxTime(r) = GetRowMinRxTime(r): mRowOriginalIndex(r) = r: mRowValidInput(r) = True
        If mIdxRxCnt > 0 Then mRowRXCount(r) = CLng(mData(r, mIdxRxCnt)) Else mRowRXCount(r) = 0
    Next r
End Sub

Private Function FindNextConflictStart() As Long
    Dim groupStart As Long, groupEnd As Long
    Dim rows() As Long, rowCount As Long

    groupStart = mScanPos
    Do While groupStart <= mFilteredCount
        groupEnd = groupStart
        Do While groupEnd < mFilteredCount
            If mCurrentSFN(groupEnd + 1) <> mCurrentSFN(groupStart) Then Exit Do
            groupEnd = groupEnd + 1
        Loop
        rowCount = groupEnd - groupStart + 1
        If rowCount > 1 Then
            rows = ExtractAbsoluteRows(groupStart, groupEnd)
            If Not ValidateBucketRows(rows, rowCount) Then
                FindNextConflictStart = groupStart
                Exit Function
            End If
        End If
        groupStart = groupEnd + 1
    Loop
End Function

Private Sub BuildPoolFromConflictStart(ByVal startRow As Long)
    Dim leftRow As Long, rightRow As Long, i As Long
    If startRow < 1 Or startRow >= mFilteredCount Then Exit Sub
    leftRow = startRow: rightRow = startRow
    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = mCurrentSFN(leftRow): leftRow = leftRow - 1: Loop
    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = mCurrentSFN(rightRow): rightRow = rightRow + 1: Loop
    Do While leftRow > 1
        If (mCurrentSFN(leftRow) - mCurrentSFN(leftRow - 1)) > 1 Then Exit Do
        leftRow = leftRow - 1
    Loop
    Do While rightRow < mFilteredCount
        If (mCurrentSFN(rightRow + 1) - mCurrentSFN(rightRow)) > 1 Then Exit Do
        rightRow = rightRow + 1
    Loop
    mPoolCount = rightRow - leftRow + 1: If mPoolCount <= 0 Then Exit Sub
    ReDim mPoolRows(1 To mPoolCount)
    For i = 1 To mPoolCount: mPoolRows(i) = leftRow + i - 1: Next i
    mPoolMinSFN = mCurrentSFN(leftRow): mPoolMaxSFN = mCurrentSFN(rightRow): mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#: If mPoolCount > mMaxObservedPoolSize Then mMaxObservedPoolSize = mPoolCount
    If DEBUG_TXSFNCR Then Debug.Print "POOL built: startRow=" & startRow & " leftRow=" & leftRow & " rightRow=" & rightRow & " poolCount=" & mPoolCount & " minSFN=" & mPoolMinSFN & " maxSFN=" & mPoolMaxSFN
    BuildPoolCests
End Sub

Private Sub BuildPoolCests()
    Dim i As Long, r As Long, startIdx As Long, curSFN As Long
    Dim rows() As Long, rowCount As Long
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
        rowCount = i - startIdx + 1
        If rowCount > 1 Then
            rows = ExtractPoolRows(startIdx, i)
            If Not ValidateBucketRows(rows, rowCount) Then
                mPoolCestCount = mPoolCestCount + 1: mPoolCestStartRows(mPoolCestCount) = startIdx: mPoolCestEndRows(mPoolCestCount) = i: mPoolCestSFN(mPoolCestCount) = curSFN
                If DEBUG_TXSFNCR Then Debug.Print "CEST chunk: idx=" & mPoolCestCount & " rows=" & startIdx & ".." & i & " sfn=" & curSFN & " count=" & rowCount
            End If
        End If
        i = i + 1
    Loop
End Sub

Private Function ResolveEntirePool() As Boolean
    Dim cestIdx As Long, rows() As Long, rowCount As Long
    Dim movedInPass As Boolean, movedAny As Boolean
    Dim passCount As Long, maxPasses As Long
    Dim order() As Long, k As Long, orderedCestIdx As Long

    If mPoolCount <= 0 Then Exit Function
    maxPasses = IIf(mPoolCount > 0, mPoolCount * 2, 1)

    Do
        BuildPoolCests
        If mPoolCestCount <= 0 Then Exit Do

        movedInPass = False
        order = BuildPoolCsetResolutionOrder()
        For k = 1 To mPoolCestCount
            orderedCestIdx = order(k)
            rowCount = mPoolCestEndRows(orderedCestIdx) - mPoolCestStartRows(orderedCestIdx) + 1
            If rowCount > 1 Then
                rows = ExtractPoolRows(mPoolCestStartRows(orderedCestIdx), mPoolCestEndRows(orderedCestIdx))
                If DEBUG_TXSFNCR Then Debug.Print "TryResolveCset: cestIdx=" & orderedCestIdx & " rowCount=" & rowCount & " sourceSFN=" & mPoolCestSFN(orderedCestIdx)
                If TryResolveCset(rows, rowCount, mPoolCestSFN(orderedCestIdx)) Then
                    movedInPass = True
                    movedAny = True
                    mPoolCountResolved = mPoolCountResolved + 1
                Else
                    mUnresolvedAttemptCount = mUnresolvedAttemptCount + 1
                End If
            End If
        Next k

        passCount = passCount + 1
        If Not movedInPass Then Exit Do
    Loop While passCount < maxPasses

    ResolveEntirePool = movedAny
End Function

Private Function TryResolveCset(ByRef rows() As Long, ByVal rowCount As Long, ByVal sourceSFN As Long) As Boolean
    Dim analysis As TCsetAnalysis: analysis = AnalyzeCsetSingleMoves(rows, rowCount)
    TryResolveCset = (analysis.candidateCount > 0 And TryPlaceOneMovedRow_NoSourceRetest(analysis.candidateRows, analysis.candidateCount, sourceSFN))
End Function

Private Function AnalyzeCsetSingleMoves(ByRef rows() As Long, ByVal rowCount As Long) As TCsetAnalysis
    Dim a As TCsetAnalysis, i As Long, j As Long, tmp As Long
    a.MovesRequired = IIf(rowCount > 1, rowCount - 1, 0): a.candidateCount = rowCount: ReDim a.candidateRows(1 To rowCount)
    For i = 1 To rowCount: a.candidateRows(i) = rows(i): Next i
    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If CompareCandidatePriority(a.candidateRows(j), a.candidateRows(i)) < 0 Then
                tmp = a.candidateRows(i)
                a.candidateRows(i) = a.candidateRows(j)
                a.candidateRows(j) = tmp
            End If
        Next j
    Next i
    AnalyzeCsetSingleMoves = a
End Function

Private Function TryPlaceOneMovedRow_NoSourceRetest(ByRef candidateRows() As Long, ByVal candidateCount As Long, ByVal sourceSFN As Long) As Boolean
    Dim i As Long, rowIdx As Long, testSFN As Long, delta As Long
    Dim originalSFN As Long, maxDec As Long, maxInc As Long
    For i = 1 To candidateCount
        rowIdx = candidateRows(i): originalSFN = mCurrentSFN(rowIdx)
        maxDec = GetMaxAllowedDecrement(rowIdx, sourceSFN)
        For delta = 1 To maxDec
            testSFN = sourceSFN - delta
            If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                mCurrentSFN(rowIdx) = testSFN
                If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                    If DEBUG_TXSFNCR Then Debug.Print "ACCEPT move: rowIdx=" & rowIdx & " sourceSFN=" & sourceSFN & " testSFN=" & testSFN & " minRx=" & mRowMinRxTime(rowIdx) & " txID=" & mRowTXID(rowIdx)
                    TryPlaceOneMovedRow_NoSourceRetest = True
                    Exit Function
                End If
                mCurrentSFN(rowIdx) = originalSFN
            End If
        Next delta

        maxInc = GetMaxAllowedIncrement(rowIdx, sourceSFN)
        For delta = 1 To maxInc
            testSFN = sourceSFN + delta
            If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                mCurrentSFN(rowIdx) = testSFN
                If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                    If DEBUG_TXSFNCR Then Debug.Print "ACCEPT move: rowIdx=" & rowIdx & " sourceSFN=" & sourceSFN & " testSFN=" & testSFN & " minRx=" & mRowMinRxTime(rowIdx) & " txID=" & mRowTXID(rowIdx)
                    TryPlaceOneMovedRow_NoSourceRetest = True
                    Exit Function
                End If
                mCurrentSFN(rowIdx) = originalSFN
            End If
        Next delta
    Next i
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If testSFN = sourceSFN Then Exit Function
    If Not IsMoveWithinRowBounds(rowIdx, sourceSFN, testSFN) Then Exit Function
    If Not IsBitmapSFNAllowed(testSFN) Then Exit Function
    If Not EvaluatePoolBucketExcludingRow(sourceSFN, rowIdx) Then Exit Function
    If Not EvaluatePoolBucketWithAddedRow(testSFN, rowIdx) Then Exit Function
    IsOneMovedRowPlacementLegal_NoSourceRetest = True
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
    Dim rows() As Long, rowCount As Long
    rows = CollectRowsForSFN(sfnVal, excludeRowIdx, -1, rowCount)
    EvaluatePoolBucketExcludingRow = ValidateBucketRows(rows, rowCount)
End Function
Private Function EvaluatePoolBucketWithAddedRow(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean
    Dim rows() As Long, rowCount As Long
    rows = CollectRowsForSFN(sfnVal, -1, addedRowIdx, rowCount)
    EvaluatePoolBucketWithAddedRow = ValidateBucketRows(rows, rowCount)
End Function
Private Function IsMoveWithinRowBounds(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    Dim rxMin As Double
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If testSFN < 0 Then Exit Function
    If testSFN < (sourceSFN - mMaxTXSFNDecrement) Then Exit Function
    If testSFN > (sourceSFN + mMaxTXSFNIncrement) Then Exit Function
    rxMin = mRowMinRxTime(rowIdx)
    If rxMin <> NO_RX_TIME Then
        If CDbl(testSFN) > rxMin Then Exit Function
    End If
    IsMoveWithinRowBounds = True
End Function
Private Function IsBitmapSFNAllowed(ByVal testSFN As Long) As Boolean
    Dim bitIdx As Long, bitChar As String
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
    Dim i As Long, rowIdx As Long
    Dim poolOrderedRows() As Long
    If mPoolCount <= 0 Then Exit Sub
    poolOrderedRows = ExtractPoolRows(1, mPoolCount)
    SortRowIndexByCurrentSFN poolOrderedRows, LBound(poolOrderedRows), UBound(poolOrderedRows)
    For i = 1 To mPoolCount
        rowIdx = poolOrderedRows(i)
        If rowIdx >= 1 And rowIdx <= mFilteredCount Then
            If Not mWritten(rowIdx) Then
                CopyRowToOutput rowIdx, mOutputWritePos
                mOutputRowSource(mOutputWritePos) = rowIdx
                mWritten(rowIdx) = True
                mOutputWritePos = mOutputWritePos + 1
            End If
        End If
    Next i
End Sub

Private Sub WriteUnwrittenRowsToOutput()
    Dim r As Long, rowCount As Long
    Dim remainingRows() As Long
    ReDim remainingRows(1 To mFilteredCount)
    For r = 1 To mFilteredCount
        If Not mWritten(r) Then
            rowCount = rowCount + 1
            remainingRows(rowCount) = r
        End If
    Next r
    If rowCount > 0 Then
        ReDim Preserve remainingRows(1 To rowCount)
        SortRowIndexByCurrentSFN remainingRows, 1, rowCount
        For r = 1 To rowCount
            CopyRowToOutput remainingRows(r), mOutputWritePos
            mOutputRowSource(mOutputWritePos) = remainingRows(r)
            mWritten(remainingRows(r)) = True
            mOutputWritePos = mOutputWritePos + 1
        Next r
    End If
    mOutputCount = mOutputWritePos - 1
End Sub

Private Sub RecomputeFinalTXperSFN()
    Dim r As Long, prevSFN As Long, txPer As Long, curSFN As Long
    Dim outputRows As Long
    If mFilteredCount <= 0 Then Exit Sub
    outputRows = mOutputWritePos - 1
    If outputRows <= 0 Then Exit Sub
    prevSFN = CLng(mOutputData(1, mIdxSFNCol))
    txPer = 0
    For r = 1 To outputRows
        curSFN = CLng(mOutputData(r, mIdxSFNCol))
        If r = 1 Then
            txPer = 1
        ElseIf curSFN = prevSFN Then
            txPer = txPer + 1
        Else
            txPer = 1
            prevSFN = curSFN
        End If
        If txPer < 1 Then txPer = 1
        If mIdxTXperSFN > 0 Then mOutputData(r, mIdxTXperSFN) = txPer
    Next r
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
Private Sub CopyRowToOutput(ByVal rowIdx As Long, ByVal outRowIdx As Long)
    Dim c As Long, colCount As Long
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Sub
    If outRowIdx < 1 Or outRowIdx > mFilteredCount Then Exit Sub
    colCount = UBound(mOutputData, 2)
    For c = LBound(mData, 2) To colCount: mOutputData(outRowIdx, c) = mData(rowIdx, c): Next c
    mOutputData(outRowIdx, mIdxSFNCol) = mCurrentSFN(rowIdx)
End Sub

Private Function BuildSubsetExcludingOne(ByRef rowListIn() As Long, ByVal rowCountIn As Long, ByVal removeRowIdx As Long, ByRef rowListOut() As Long) As Long: BuildSubsetExcludingOne = 0: End Function
Private Sub Sort3RowsByMinRxTime(ByRef rowA As Long, ByRef rowB As Long, ByRef rowC As Long): End Sub
Private Sub QuickSortLongs(ByRef arr() As Long, ByVal first As Long, ByVal last As Long): End Sub
Private Sub SortRowIndexByCurrentSFN(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As Long, tmp As Long
    i = first
    j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While CompareRowOrder(arr(i), pivot) < 0: i = i + 1: Loop
        Do While CompareRowOrder(arr(j), pivot) > 0: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then SortRowIndexByCurrentSFN arr, first, j
    If i < last Then SortRowIndexByCurrentSFN arr, i, last
End Sub
Private Function CompareRowOrder(ByVal rowA As Long, ByVal rowB As Long) As Long
    CompareRowOrder = Sgn(mCurrentSFN(rowA) - mCurrentSFN(rowB))
    If CompareRowOrder = 0 Then CompareRowOrder = Sgn(mRowOriginalIndex(rowA) - mRowOriginalIndex(rowB))
End Function
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
    ReDim mOutputRowSource(1 To mFilteredCount)
    mOutputCount = 0
    mOutputWritePos = 1
End Sub
Private Function ExtractAbsoluteRows(ByVal startRow As Long, ByVal endRow As Long) As Long()
    Dim rows() As Long, i As Long, n As Long
    n = endRow - startRow + 1
    ReDim rows(1 To n)
    For i = 1 To n: rows(i) = startRow + i - 1: Next i
    ExtractAbsoluteRows = rows
End Function
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
    If rowCount <= 1 Then
        ValidateBucketRows = True
        Exit Function
    End If
    If Not IsGroupTXIDUnique(rows, rowCount) Then Exit Function
    If Not IsGroupMergedStationsUnique(rows, rowCount) Then Exit Function
    ValidateBucketRows = True
End Function

Private Function IsGroupTXIDUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim dictTx As Object, i As Long, rowIdx As Long, txIDKey As String
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
    Dim dictStations As Object, i As Long, rowIdx As Long
    Set dictStations = CreateObject("Scripting.Dictionary")
    For i = 1 To rowCount
        rowIdx = rows(i)
        If Not AddRowStationsToDict(rowIdx, dictStations) Then Exit Function
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
    Dim colIdx As Long, dataUB As Long
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
    If Not EvaluatePoolBucketWithAddedRow(movedSFN, movedRowIdx) Then Exit Function
    DoesMovedRowFormValidLocalGroup = True
End Function

Private Function GetMaxAllowedDecrement(ByVal rowIdx As Long, ByVal sourceSFN As Long) As Long
    Dim maxDec As Long
    maxDec = mMaxTXSFNDecrement
    If sourceSFN < maxDec Then maxDec = sourceSFN
    If maxDec < 0 Then maxDec = 0
    GetMaxAllowedDecrement = maxDec
End Function

Private Function GetMaxAllowedIncrement(ByVal rowIdx As Long, ByVal sourceSFN As Long) As Long
    Dim maxInc As Long
    maxInc = mMaxTXSFNIncrement
    If maxInc < 0 Then maxInc = 0
    GetMaxAllowedIncrement = maxInc
End Function

Private Function CompareCandidatePriority(ByVal rowA As Long, ByVal rowB As Long) As Long
    If mRowRXCount(rowA) <> mRowRXCount(rowB) Then
        CompareCandidatePriority = Sgn(mRowRXCount(rowB) - mRowRXCount(rowA))
    ElseIf mRowMinRxTime(rowA) <> mRowMinRxTime(rowB) Then
        CompareCandidatePriority = Sgn(mRowMinRxTime(rowA) - mRowMinRxTime(rowB))
    Else
        CompareCandidatePriority = Sgn(rowA - rowB)
    End If
End Function

Private Function BuildPoolCsetResolutionOrder() As Long()
    Dim order() As Long
    Dim i As Long, j As Long, tmp As Long
    ReDim order(1 To mPoolCestCount)
    For i = 1 To mPoolCestCount: order(i) = i: Next i
    For i = 1 To mPoolCestCount - 1
        For j = i + 1 To mPoolCestCount
            If CompareCsetOrder(order(j), order(i)) < 0 Then
                tmp = order(i)
                order(i) = order(j)
                order(j) = tmp
            End If
        Next j
    Next i
    BuildPoolCsetResolutionOrder = order
End Function

Private Function CompareCsetOrder(ByVal csetA As Long, ByVal csetB As Long) As Long
    Dim distA As Double, distB As Double
    distA = Abs(CDbl(mPoolCestSFN(csetA)) - mPoolCenter)
    distB = Abs(CDbl(mPoolCestSFN(csetB)) - mPoolCenter)
    If distA <> distB Then
        CompareCsetOrder = Sgn(distA - distB)
    ElseIf mPoolCestSFN(csetA) <> mPoolCestSFN(csetB) Then
        CompareCsetOrder = Sgn(mPoolCestSFN(csetA) - mPoolCestSFN(csetB))
    Else
        CompareCsetOrder = Sgn(csetA - csetB)
    End If
End Function

Private Sub LoadDirectionalMoveBounds()
    Dim fallbackDelta As Long
    fallbackDelta = ResolveNamedMoveBound("maxTX_SFNdelta", 1)
    mMaxTXSFNDecrement = ResolveNamedMoveBound("maxTX_SFN_est_decrement", fallbackDelta)
    mMaxTXSFNIncrement = ResolveNamedMoveBound("maxTX_SFN_est_increment", fallbackDelta)
End Sub

Private Function ResolveNamedMoveBound(ByVal nameText As String, ByVal defaultVal As Long) As Long
    Dim nm As Name, v As Variant
    On Error GoTo UseDefault
    Set nm = ThisWorkbook.Names(nameText)
    v = Application.Evaluate(nm.RefersTo)
    If IsError(v) Or Not IsNumeric(v) Then GoTo UseDefault
    ResolveNamedMoveBound = CLng(CDbl(v))
    If ResolveNamedMoveBound < 0 Then ResolveNamedMoveBound = 0
    Exit Function
UseDefault:
    ResolveNamedMoveBound = defaultVal
End Function

Private Function GetRowMinRxTime(ByVal rowIdx As Long) As Double
    Dim st As Long, colIdx As Long, rxVal As Double, dataUB As Long
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
