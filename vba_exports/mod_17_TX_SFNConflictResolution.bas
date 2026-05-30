Attribute VB_Name = "mod_17_TX_SFNConflictResolution"
Option Explicit

' Version: V1.0.0
Private Const MODULE_VERSION_TXSFNCR As String = "V1.0.0"

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

    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, txBitmap, bitmapLen
    If Not ValidateInputMonotoneTXSFN() Then Exit Sub
    PrepareRowDerivedData
    InitializeOutputBuffer

    Do
        t0 = MicroTimer_TXSFNCR()
        conflictStart = FindNextConflictStart()
        If conflictStart <= 0 Then Exit Do
        BuildPoolFromConflictStart conflictStart
        ResolveEntirePool
        WriteResolvedPoolToOutput
        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR() - t0)
    Loop

    t0 = MicroTimer_TXSFNCR()
    WriteUnwrittenRowsToOutput
    RecomputeFinalTXperSFN
    FinalizeOutputVariant data
    mFinalizeSeconds = mFinalizeSeconds + (MicroTimer_TXSFNCR() - t0)
    elapsedSeconds = MicroTimer_TXSFNCR() - startTime
End Sub

Private Sub InitializeContext(ByRef data As Variant, ByVal filteredCount As Long, ByVal idxSFNCol As Long, ByVal idxTXID As Long, ByVal idxTXQ As Long, ByVal idxLEN As Long, ByVal idxTXperSFN As Long, ByVal idxRxCnt As Long, ByVal idxAvg As Long, ByVal idxTotLat As Long, ByVal idxGen As Long, ByRef rxDataColIdx() As Long, ByRef rxStationIDs() As Long, ByVal activeRxCount As Long, ByRef dictS2V As Object, ByRef dictVC As Object, ByRef dictA2P As Object, ByRef dictP2R As Object, ByRef dictP2Sigma As Object, ByVal txBitmap As String, ByVal bitmapLen As Long)
    mData = data: mFilteredCount = filteredCount: mIdxSFNCol = idxSFNCol: mIdxTXID = idxTXID: mIdxTXQ = idxTXQ: mIdxLEN = idxLEN: mIdxTXperSFN = idxTXperSFN: mIdxRxCnt = idxRxCnt: mIdxAvg = idxAvg: mIdxTotLat = idxTotLat: mIdxGen = idxGen
    mRxDataColIdx = rxDataColIdx: mRxStationIDs = rxStationIDs: mActiveRxCount = activeRxCount
    Set mDictS2V = dictS2V: Set mDictVC = dictVC: Set mDictA2P = dictA2P: Set mDictP2R = dictP2R: Set mDictP2Sigma = dictP2Sigma
    mTxBitmap = txBitmap: mBitmapLen = bitmapLen: mScanPos = 1: mOutputWritePos = 1: mPoolCount = 0: mPoolCestCount = 0
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
        mInitialSFN(r) = CLng(mData(r, mIdxSFNCol)): mCurrentSFN(r) = mInitialSFN(r): mRowTXID(r) = CLng(mData(r, mIdxTXID)): mRowTXQTime(r) = CDbl(mData(r, mIdxTXQ)): mRowNsch(r) = CLng(mData(r, mIdxLEN)): mRowPduKey(r) = CStr(mData(r, mIdxLEN)): mRowMinRxTime(r) = CDbl(mData(r, mIdxTXQ)): mRowOriginalIndex(r) = r: mRowValidInput(r) = True
        If mIdxRxCnt > 0 Then mRowRXCount(r) = CLng(mData(r, mIdxRxCnt)) Else mRowRXCount(r) = 0
    Next r
End Sub

Private Function FindNextConflictStart() As Long
    Dim r As Long
    For r = mScanPos To mFilteredCount - 1
        If mCurrentSFN(r) = mCurrentSFN(r + 1) Then FindNextConflictStart = r: Exit Function
    Next r
End Function

Private Sub BuildPoolFromConflictStart(ByVal startRow As Long)
    Dim leftRow As Long, rightRow As Long, i As Long
    If startRow < 1 Or startRow >= mFilteredCount Then Exit Sub
    leftRow = startRow: rightRow = startRow + 1
    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = mCurrentSFN(leftRow): leftRow = leftRow - 1: Loop
    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = mCurrentSFN(rightRow): rightRow = rightRow + 1: Loop
    mPoolCount = rightRow - leftRow + 1: If mPoolCount <= 0 Then Exit Sub
    ReDim mPoolRows(1 To mPoolCount)
    For i = 1 To mPoolCount: mPoolRows(i) = leftRow + i - 1: Next i
    mPoolMinSFN = mCurrentSFN(leftRow): mPoolMaxSFN = mCurrentSFN(rightRow): mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#: If mPoolCount > mMaxObservedPoolSize Then mMaxObservedPoolSize = mPoolCount
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
        i = i + 1
    Loop
End Sub

Private Sub ResolveEntirePool()
    Dim cestIdx As Long, rows() As Long, rowCount As Long
    If mPoolCestCount <= 0 Then Exit Sub
    For cestIdx = 1 To mPoolCestCount
        rowCount = mPoolCestEndRows(cestIdx) - mPoolCestStartRows(cestIdx) + 1
        If rowCount > 0 Then rows = ExtractPoolRows(mPoolCestStartRows(cestIdx), mPoolCestEndRows(cestIdx)): Call TryResolveCset(rows, rowCount, mPoolCestSFN(cestIdx))
    Next cestIdx
End Sub

Private Function TryResolveCset(ByRef rows() As Long, ByVal rowCount As Long, ByVal sourceSFN As Long) As Boolean
    Dim analysis As TCsetAnalysis: analysis = AnalyzeCsetSingleMoves(rows, rowCount)
    TryResolveCset = (analysis.candidateCount > 0 And TryPlaceOneMovedRow_NoSourceRetest(analysis.candidateRows, analysis.candidateCount, sourceSFN))
End Function

Private Function AnalyzeCsetSingleMoves(ByRef rows() As Long, ByVal rowCount As Long) As TCsetAnalysis
    Dim a As TCsetAnalysis, i As Long
    a.MovesRequired = IIf(rowCount > 1, rowCount - 1, 0): a.candidateCount = rowCount: ReDim a.candidateRows(1 To rowCount)
    For i = 1 To rowCount: a.candidateRows(i) = rows(i): Next i
    AnalyzeCsetSingleMoves = a
End Function

Private Function TryPlaceOneMovedRow_NoSourceRetest(ByRef candidateRows() As Long, ByVal candidateCount As Long, ByVal sourceSFN As Long) As Boolean
    Dim i As Long, rowIdx As Long, testSFN As Long
    For i = 1 To candidateCount
        rowIdx = candidateRows(i): testSFN = sourceSFN + 1
        If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then mCurrentSFN(rowIdx) = testSFN: mWritten(rowIdx) = True: TryPlaceOneMovedRow_NoSourceRetest = True: Exit Function
    Next i
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If testSFN < sourceSFN Then Exit Function
    If Not IsMoveWithinRowBounds(rowIdx, testSFN) Then Exit Function
    If Not IsBitmapSFNAllowed(testSFN) Then Exit Function
    IsOneMovedRowPlacementLegal_NoSourceRetest = True
End Function

Private Function TryForwardEscapeMove_NoSourceRetest(ByVal sourceSFN As Long, ByVal rowIdx As Long) As Boolean: TryForwardEscapeMove_NoSourceRetest = False: End Function
Private Function ResolveTripleSplitDeterministic(ByRef rows() As Long, ByVal sourceSFN As Long) As Boolean
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN + 1) Then mCurrentSFN(rows(i)) = sourceSFN + 1: ResolveTripleSplitDeterministic = True: Exit Function
    Next i
End Function
Private Function IsPoolForcedSplitFirstMoveLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean: IsPoolForcedSplitFirstMoveLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN): End Function
Private Function IsPoolSingleRowPlacementLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean: IsPoolSingleRowPlacementLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN): End Function
Private Function EvaluatePoolBucketExcludingRow(ByVal sfnVal As Long, ByVal excludeRowIdx As Long) As Boolean: EvaluatePoolBucketExcludingRow = True: End Function
Private Function EvaluatePoolBucketWithAddedRow(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean: EvaluatePoolBucketWithAddedRow = True: End Function
Private Function IsMoveWithinRowBounds(ByVal rowIdx As Long, ByVal testSFN As Long) As Boolean: IsMoveWithinRowBounds = True: End Function
Private Function IsBitmapSFNAllowed(ByVal testSFN As Long) As Boolean: IsBitmapSFNAllowed = True: End Function

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
    Dim r As Long, prevSFN As Long, txPer As Long
    If mFilteredCount <= 0 Then Exit Sub
    prevSFN = mCurrentSFN(1): txPer = 1
    For r = 1 To mFilteredCount
        If r = 1 Then
            txPer = 1
        ElseIf mCurrentSFN(r) = prevSFN Then
            txPer = txPer + 1
        Else
            txPer = 1
            prevSFN = mCurrentSFN(r)
        End If
        If mIdxTXperSFN > 0 Then mOutputData(r, mIdxTXperSFN) = txPer
    Next r
End Sub

Private Sub FinalizeOutputVariant(ByRef data As Variant): data = mOutputData: End Sub
Private Sub CopyRowToOutput(ByVal rowIdx As Long)
    Dim c As Long, colCount As Long
    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Sub
    colCount = UBound(mOutputData, 2)
    For c = LBound(mData, 2) To colCount: mOutputData(rowIdx, c) = mData(rowIdx, c): Next c
    mOutputData(rowIdx, mIdxSFNCol) = mCurrentSFN(rowIdx)
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
Private Function SafeDiv(ByVal numerator As Double, ByVal denominator As Double) As Double: If denominator = 0# Then SafeDiv = 0# Else SafeDiv = numerator / denominator: End If: End Function
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
